# Railway Imprest Bill Automator

A Streamlit app that automates generating the two monthly imprest-account
Word documents for **SETTIHALLI Station** (Indian Railways):

- **Annexure I** — Master Bill (recoupment, balances, voucher summary, certification)
- **Annexure II** — Labour Annexure (attendance & wage register for two cleaning workers)

Instead of manually filling Word templates every month, you enter the
month, expenses, vouchers, and worker attendance in a browser UI, and the
app fills the existing `.docx` templates and gives you both files to
download.

---

## Project structure

```
railway_imprest/
├── app.py                     # entrypoint — wires everything together
├── config.py                  # MONTHS, SANCTIONED_AMOUNT, DAILY_WAGE, template paths
├── requirements.txt
├── App.bat                    # Windows one-click launcher
├── imprest_balances.ini       # auto-generated at runtime — do not hand-edit
│
├── templates/                 # your 5 fixed .docx templates go here (see below)
│
├── core/
│   ├── dates.py                # get_days_in_month()
│   ├── num_to_words.py         # number → Indian-English words (lakh/thousand)
│   ├── state.py                 # shared session-state default-init helper
│   ├── validation.py            # template anchor-text sanity checks
│   └── balance_store.py         # month-to-month closing-balance persistence (INI)
│
├── docx_engine/
│   ├── xml_helpers.py           # raw OOXML run/cell formatting helpers
│   ├── annexure1.py              # generate_annexure_1()
│   └── labour.py                 # generate_labour_annexure()
│
└── ui/
    ├── styles.py                 # global CSS (inject_css())
    ├── sidebar.py                 # month/year selector
    ├── tab_expenses.py            # Tab 1 — Consumables
    ├── tab_labour.py              # Tab 2 — Labour
    └── tab_generate.py            # Tab 3 — balances, generate, download
```

---

## Setup

### 1. Add your templates

Place your 5 existing Word templates in `templates/`, with these **exact**
filenames (note: `ANNEXTURE` for the labour files, spelled as in the
original folder):

```
templates/
├── ANNEXURE 1 WITH 3 VOUCHERS.docx
├── ANNEXURE 1 WITH 4 VOUCHERS With Stationery.docx
├── ANNEXURE 1 WITH 4 VOUCHERS Without Stationery.docx
├── ANNEXTURE II LABOUR 3 VOUCHERS.docx
└── ANNEXTURE II LABOUR 4 VOUCHERS.docx
```

> Templates are **not** included in this repo — add them yourself locally
> (or via Git LFS / a private release asset if you want them versioned).

### 2. Install dependencies

```bash
python -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Run

**Windows:** double-click `App.bat` — it checks that the app files and all
5 templates are present, installs missing dependencies if needed, then
launches the app.

**Manual (any OS):**
```bash
streamlit run app.py
```

Run this from inside the `railway_imprest/` folder — the internal module
imports and the `templates/`/`imprest_balances.ini` paths are relative to it.

The app opens at `http://localhost:8501`.

---

## Features

- **Two voucher modes** — 3-voucher (single consumable bill) or 4-voucher
  (with/without stationery), each mapped to its own fixed template pair.
- **Auto-carried opening balance** — after generating a month's documents,
  its closing balance is saved to `imprest_balances.ini`; the next month's
  Opening Balance field auto-fills from it (editable if needed).
- **Per-month input memory** — amounts, dates, and worked-days entered for
  one month don't bleed into another month; switching months shows that
  month's own saved values (or a fresh default if untouched).
- **Stale-document protection** — if you change any input after
  generating, the previously generated files are cleared automatically so
  you can't accidentally download an out-of-date document.
- **Template sanity check** — before generating, each template is opened
  and checked for the expected tables/anchor text, with a clear error if
  a template was edited and no longer matches what the app expects.
- **Negative-balance / leave-overlap guard** — the Generate button is
  disabled if the closing balance would go negative or leave days
  overlap between the two workers.

---

## Notes

- `imprest_balances.ini` persists on the machine/deployment running the
  app. If you deploy somewhere with an ephemeral filesystem (some cloud
  hosts wipe local disk on redeploy), this file — and the auto-carried
  opening balance — won't survive a redeploy. Consider excluding it from
  `.gitignore` if you want to seed history, or move to a hosted DB if you
  need it to survive redeploys.
- This app is built for a single station (SETTIHALLI) with fixed daily
  wage/sanctioned-amount constants in `config.py` — edit those directly
  if either changes.

---

## Suggested `.gitignore`

```
venv/
__pycache__/
*.pyc
imprest_balances.ini
templates/*.docx
```
