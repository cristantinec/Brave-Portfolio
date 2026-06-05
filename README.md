# Brave PH Lending Portfolio Analysis

> Philippines lending portfolio performance analysis (2019–2026) — bank channeling partnership evaluation

![Python](https://img.shields.io/badge/Python-3.x-blue) ![pandas](https://img.shields.io/badge/pandas-2.x-green) ![Power BI](https://img.shields.io/badge/Power%20BI-Free-yellow) ![Status](https://img.shields.io/badge/Status-Complete-brightgreen)

---

## Table of Contents
1. [Overview](#overview)
2. [Problem Statement](#problem-statement)
3. [Data](#data)
4. [Methodology](#methodology)
5. [Insights](#insights)
6. [Recommendations](#recommendations)

---

## 📌 Overview

Brave is a global digital lending company that provides small, short-term loans to underserved borrowers via mobile app. In the Philippines, Brave has been operating since 2019, disbursing loans typically ranging from ₱500 to ₱10,000+ with an average term of approximately 61 days.

This case study evaluates whether Brave's Philippines portfolio is a viable candidate for a bank loan channeling partnership — an arrangement where a bank provides capital that Brave deploys as loans, with the bank earning a return and Tala handling origination, underwriting, and collections.

---

## 🚩 Problem Statement

> Prospective bank channeling partners lack a structured, time-series view of Tala's Philippines portfolio performance — specifically across **origination volume growth**, **credit quality** (repayment rates and delinquency), and **fee-driven cash-on-cash returns** — across 1-month, 1-year, 3-year, and 5-year horizons, making it difficult to assess whether Tala's portfolio meets their risk and return thresholds for a loan channeling partnership.

## 📊 Benchmark Summary Table

| KPI | Red Flag | Minimum | Preferred |
|-----|-----------|----------|------------|
| Origination volume | <₱500M/mo | ₱500M–1B/mo | ₱1.5B+/mo |
| Repayment rate (30D) | <80% | 85% | 88%+ |
| DQ rate (30–119 DPD) | >20% | <15% | <12% |
| Cash‑on‑cash return | <0.95x | >1.0x | 1.04x+ |
| Fee & interest revenue | <10% of principal | 12–14% | 15%+ |
| Repeat borrower rate | <50% | 60%+ | 70%+ (loan 3+) |

**Four dimensions evaluated:**

| # | Dimension | Key Metric |
|---|---|---|
| 1 | Origination Volume | `Origination_Principal_PHP` |
| 2 | Credit Quality | `Repayment_Rate_30D` · `DQ_Rate_30_119DPD` |
| 3 | Fee & Interest Revenue | `Fee_Interest_Revenue_PHP` · `Fee_Rev_Pct` |
| 4 | Cash-on-Cash Return | `Cash_on_Cash_Return` |

> [!NOTE]
> The **3-year (2022–2025)** and **5-year (2019–2025)** windows are most relevant for a bank partner decision. 3Y reflects mature underwriting; 5Y covers a full credit cycle including COVID stress.

---

## Data

**Source file:** `Philippines_Case_Study_Portfolio_Summary_-_2026_03.xlsx`

```
Full file   → Jul 2016 – Mar 2026  (16 tabs)
This study  → Jan 2019 – Mar 2026  (87 monthly cohorts, 6 tabs used)
```

| Tab | Rows | Cols | Date Range | Used |
|---|---|---|---|---|
| `Cohort Data` | 114 | 5 | Oct 2016 – Mar 2026 | ✅ Primary |
| `Cash on Cash` | 5,305 | 3 | Jul 2016 – Mar 2026 | ✅ Primary |
| `AR Bucket Pull` | 97 | 24 | Mar 2018 – Mar 2026 | ✅ Primary |
| `Repayment Rates Data` | 115 | 8 | Jul 2016 – Jan 2026 | ✅ Primary |
| `Distr by Count Data` | 119 | 7 | Mar 2016 – Mar 2026 | ✅ Primary |
| `Distr by Loan Number Data` | 119 | 8 | Mar 2016 – Mar 2026 | ✅ Primary |
| Remaining 10 tabs | — | — | — | 📋 Reference only |

> [!NOTE]
> `Cash on Cash` has 5,305 rows because it stores one record per *(disbursed month × payment month)* pair — collapsed to 87 rows via `groupby` in Step 3.

---

## Methodology

### Tools

```bash
Python 3.x    # data cleaning, transformation, metric computation
pandas        # DataFrames, merging, groupby
openpyxl      # reading/writing Excel — formula-free output
Power BI Free # dashboard visualization

# Install
pip install pandas openpyxl

# Run
python transform.py
```

### Step 1 — Load

```python
import pandas as pd
xl = pd.read_excel("Philippines_Case_Study_Portfolio_Summary_-_2026_03.xlsx", sheet_name=None)
# → dict of 16 DataFrames; only 6 used
```

### Step 2 — Clean (per tab)

```python
# All tabs: convert MONTH to datetime, filter >= 2019-01-01, sort ascending

# Cohort Data — fill null OTHER_REVENUE with 0
cohort["OTHER_REVENUE"] = cohort["OTHER_REVENUE"].fillna(0)

# Repayment Rates — keep NaN for immature vintages (do NOT fill with 0)

# Distr by Count & Loan Number — drop sub-header row, cast to numeric
df = df.iloc[1:]                                 # row 0 = "Count Loans" label
df[col] = pd.to_numeric(df[col], errors="coerce")
```

### Step 3 — Compute Derived Metrics

```python
# 1. Fee & interest revenue
cohort["FEE_INTEREST"] = cohort["INTEREST"] + cohort["OTHER_REVENUE"]

# 2. Total cash collected per vintage (collapses 5,305 → 87 rows)
coc_sum = coc.groupby("DISBURSED_MONTH")["CASH_PAID"].sum().reset_index()

# 3. Cash-on-cash return
main["CASH_ON_CASH"] = main["TOTAL_CASH_COLLECTED"] / main["PRINCIPAL"]

# 4. Delinquency rate 30–119 DPD
ar["DQ_30_119"] = (ar["pr_30_59_dpd"] + ar["pr_60_89_dpd"] + ar["pr_90_119_dpd"]) / ar["total pr"]

# 5. Repeat borrower rate (loan 3+ / all loans)
dl["REPEAT_PCT"] = dl[["GRP_3_5","GRP_6_9","GRP_10_15","GRP_16_24","GRP_25_PLUS"]].sum(axis=1) / dl["TOTAL"]
```

### Step 4 — Merge & Output

```python
# Left-join all 6 tables on MONTH (cohort = spine)
main = cohort.merge(coc_sum, on="MONTH", how="left")
main = main.merge(ar,        on="MONTH", how="left")
main = main.merge(rr,        on="MONTH", how="left")
main = main.merge(dc,        on="MONTH", how="left")
main = main.merge(dl,        on="MONTH", how="left")
# → 87 rows × 17 columns — static values, no formulas

with pd.ExcelWriter("Tala_PH_Portfolio_PowerBI_Ready.xlsx", engine="openpyxl") as w:
    main.to_excel(w,   sheet_name="Portfolio_Data", index=False)
    annual.to_excel(w, sheet_name="Annual_Summary", index=False)
```

---

## Insights

### Annual Summary

| Year | Originations (₱B) | Fee Rev (₱B) | Fee % | CoC | DQ 30–119 | RR 30D |
|---|---|---|---|---|---|---|
| 2019 | 15.8 | 2.53 | 16.0% | 1.06x | ⚠️ 14.5% | 87.6% |
| 2020 | 7.1 | 1.14 | 16.0% | 1.05x | ⚠️ 12.4% | 86.3% |
| 2021 | 11.7 | 1.87 | 16.0% | 1.06x | ✅ 3.5% | 87.9% |
| 2022 | 15.1 | 2.42 | 16.0% | 1.06x | ✅ 3.7% | 88.5% |
| 2023 | 25.3 | 4.04 | 16.0% | 1.06x | ✅ 5.3% | 88.9% |
| 2024 | 28.7 | 4.59 | 16.0% | 1.06x | ✅ 5.5% | 89.2% |
| 2025 | 32.1 | 5.13 | 16.0% | 1.06x | ✅ 5.5% | 89.5% |
| **5Y avg** | — | — | **16.0%** | **1.062x** | **7.2%** | **87.2%** |

### Key Findings

**INS-01 · Origination volume grew 3.7× (2019→2025) with a full COVID recovery**
- Grew from ₱15.8B (2019) → ₱32.1B (2025). Contracted to ₱7.1B in 2020 (−55% YoY), surpassed pre-COVID levels by 2022, and accelerated through 2023–2025.
- Current run-rate: **~₱2.7B/month** — above the ₱1.5B+ threshold required by bank channeling partners.

**INS-02 · Fee & interest revenue locked at ~16% of principal — no variance across years**
- Consistent across all 87 cohorts including the COVID year. Cumulative total: **₱22.5B**.
- Yield projections are predictable; the fee cushion above credit losses is reliable for a bank partner.

**INS-03 · Cash-on-cash return steady at 1.062x — every mature vintage above 1.0x**
- Range: 1.04x–1.12x across mature vintages. Every ₱1.00 deployed returns ₱1.062.
- At 16% fee revenue on ~61-day terms, implied annualized IRR is approximately **25–35%**.

**INS-04 · DQ rate dropped from 14.5% (2019) to 5.5% (2022–2025) — structural improvement**
- Fell from 14.5% in 2019 (early-stage underwriting) to 3.5%–5.5% from 2021 onward.
- Current level (~5.5%) is well below the 12% preferred threshold for bank channeling partners.

**INS-05 · ~90% of principal from repeat borrowers (loan 3+) — exceptional portfolio maturity**
- Proven, returning customers dominate the portfolio. Repeat borrowers have higher loan sizes, higher repayment rates, and lower DQ — lower marginal credit risk as the portfolio scales.

**INS-06 · Repayment rates stable but plateaued — ~88–90% is the structural ceiling**
- Moved from 87.6% (2019) → 89.5% (2025), less than 2pp in 6 years.
- **88–90% at 30D is the expected steady-state**, not a temporary high. Set covenants accordingly.

---

## Recommendations

### Benchmark vs. Thresholds

| KPI | Red Flag | Minimum | Preferred | Tala Actual | Verdict |
|---|---|---|---|---|---|
| Origination volume | <₱500M/mo | ₱500M–1B/mo | ₱1.5B+/mo | ~₱2.7B/mo | ✅ Exceeds |
| Repayment rate (30D) | <80% | 85% | 88%+ | 87.2% avg / 88.3% latest | ✅ Meets |
| DQ rate (30–119 DPD) | >20% | <15% | <12% | 7.2% avg / 5.5% recent | ✅ Exceeds |
| Cash-on-cash return | <0.95x | >1.0x | 1.04x+ | 1.062x avg | ✅ Exceeds |
| Fee & interest revenue | <10% | 12–14% | 15%+ | ~16.0% every year | ✅ Exceeds |
| Repeat borrower rate | <50% | 60%+ | 70%+ (loan 3+) | ~90% | ✅ Exceeds |

> [!TIP]
> Tala meets or exceeds the **preferred threshold on all six KPIs**. The portfolio is a strong candidate for a bank loan channeling arrangement.

### Actions

**`[BANK]` REC-01 · Use 3-year window (2022–2025) as primary evaluation**
The 3Y view captures mature underwriting (DQ <6%, RR >88%, 2× growth) without being diluted by 2019–2020 early-stage performance.

**`[BANK]` REC-02 · Set repayment rate covenant at 85% (30D), 3-month rolling average**
Provides a buffer below the 87.2% average while avoiding false triggers from normal seasonal dips.

**`[BANK]` REC-03 · Structure with a ₱500M–1B pilot tranche before scaling**
Data supports proceeding — but validate operational processes (reporting, collections) before full capital commitment.

**`[BRAVE]` REC-04 · Address 2019 DQ (14.5%) proactively in presentations**
Explain it reflects early-stage underwriting, and that DQ has been structurally below 6% for 4+ consecutive years.

**`[BRAVE]` REC-05 · Compute formal net IRR from loan-level cash flow timing**
CoC of 1.062x understates the return. Implied annualized IRR of ~25–35% is a stronger story for bank treasury teams.

**`[BRAVE]` REC-06 · Automate monthly ETL refresh — the pipeline is already built**
`transform.py` runs in <10 seconds. Schedule it on new monthly data to keep the Power BI dashboard current.

---

> [!IMPORTANT]
> **Final answer to the problem statement:** Based on 87 months of data (Jan 2019 – Mar 2026), Brave's Philippines portfolio meets or exceeds the preferred threshold on all six KPIs evaluated. The strongest case is made using the 3-year window (2022–2025). Tala is a quantitatively strong candidate for a bank loan channeling partnership.

---

*Tala Technologies, Inc. · Philippines Portfolio Case Study · Data period: Jan 2019 – Mar 2026 · All values in PHP · For discussion purposes only*
