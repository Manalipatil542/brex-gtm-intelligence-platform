# Brex GTM Intelligence Platform
### A self-initiated B2B analytics project modeling Brex's Go-To-Market operations end-to-end

---

## What is Brex?

Brex is a fintech company that provides corporate credit cards and financial software to startups and enterprises. Unlike traditional banks, Brex evaluates creditworthiness based on company funding and cash balance — not personal credit scores. Their customers are fast-growing B2B companies.

---

## The Business Problem

As Brex scales, their Go-To-Market (GTM) team faces a critical challenge — they generate leads from 7 different sources (outbound sales, inbound demos, paid ads, partner referrals, events, SEO) but have **no unified view** of which sources are actually driving revenue.

This creates 3 specific problems:

**1. Budget Waste**
Money is being spent on Paid Google and LinkedIn ads that convert at only 7–9%, while high-converting channels like Inbound Demo (32%) and Partner Referral (28%) are under-resourced.

**2. No Lead Prioritization**
Sales reps treat all leads equally, wasting time on cold leads instead of routing hot leads instantly to Account Executives.

**3. Partner Blindspot**
Partner referrals drive the highest ARR ($409K out of $1.06M total) but there is no structured tracker to monitor partner health, engagement, and co-marketing activity.

---

## What This Project Solves

This project builds a **GTM Intelligence Platform** that gives Brex's RevOps and BizOps teams a single source of truth across the entire funnel.

---

## Project Structure

```
brex-gtm-intelligence-platform/
├── generate_data.py        # Module 1 — Synthetic lead generation
├── build_excel.py          # Module 2 — Excel intelligence model builder
├── leads.csv               # Generated dataset (500 leads)
├── summary.json            # Aggregated funnel metrics
└── Brex_GTM_Model.xlsx     # Final Excel output (open directly)
```

---

## Module Breakdown

### Module 1 — Data Generation (`generate_data.py`)
Simulates 500 realistic B2B leads across 7 acquisition sources with conversion probabilities based on real B2B benchmarks.

Each lead flows through a 4-stage funnel:
```
Lead → MQL → SQL → Opportunity → Closed Won
```
Revenue is calculated based on company segment:
| Segment | Base ARR |
|---|---|
| Startup (<50 emp) | $8,000 |
| Mid-Market (50–500) | $35,000 |
| Enterprise (500+) | $120,000 |

---

### Module 2 — Excel Intelligence Model (`build_excel.py`)
Builds a 4-sheet finance-grade Excel workbook using Python and OpenPyXL.

**Sheet 1 — Executive Summary**
KPI scorecards (500 leads, $1.06M ARR, 4.8% win rate) and full source performance breakdown table.

**Sheet 2 — Pipeline Trend**
12-month lead and ARR growth showing business momentum across the full year.

**Sheet 3 — Partner Tracker**
7 Brex partners with health scores, engagement metrics, win rates, NRR%, QBR status, and co-marketing activity.

**Sheet 4 — Lead Scoring Model**
Firmographic + behavioral scoring framework with routing logic:
| Score | Classification | Action |
|---|---|---|
| 80–100 | Hot MQL | Route instantly to AE |
| 50–79 | MQL | Enroll in SDR outreach cadence |
| 30–49 | Nurture | Marketo 30-day drip program |
| 0–29 | Cold | Monitor only — no outreach |

---

## Key Findings

- **Inbound Demo** converts at 32% — 4.5x better than Paid Google (7%)
- **Partner Referral** drives $409K ARR — the single highest revenue channel
- **Enterprise deals** average $120K ARR vs $8K for Startups
- **Zip Procurement and Wiz Security** are at-risk partners needing immediate QBR
- Reallocating 20% of paid ad budget to partner development could increase ARR by an estimated 15–20%

---

## Tools Used

| Tool | Purpose |
|---|---|
| Python | Data generation, Excel automation |
| Pandas | Data manipulation and aggregation |
| NumPy | Statistical simulation |
| OpenPyXL | Excel workbook creation and formatting |
| JSON | Summary metrics storage |
| Excel | Final deliverable |

---

## How to Run

```bash
# Step 1 — Install dependencies
pip install pandas numpy openpyxl

# Step 2 — Generate the data
python generate_data.py

# Step 3 — Build the Excel model
python build_excel.py

# Output: Brex_GTM_Model.xlsx
```

---

## Why I Built This

This project was self-initiated to demonstrate end-to-end GTM analytics thinking — from raw data generation to executive-ready reporting. It shows the ability to think like a RevOps analyst, not just run queries.

Built as part of a portfolio targeting Data Analyst, BizOps, and Revenue Operations roles.

---
