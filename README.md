# Cluster Energy Billing

Turn your monthly **meter readings** and **electricity supplier invoices** into a
clear, itemised **PDF energy invoice for every household** in a shared solar +
battery microgrid ("cluster"), plus the **CSV files needed to record those
invoices in Xero**.

One command does the whole job:

```bash
python3 cluster_energy_billing.py --months Apr-Jun \
  --meters data/meters_2026-03.csv data/meters_2026-04.csv data/meters_2026-05.csv data/meters_2026-06.csv \
  --pdfs data/supplier_invoice_2026-04.pdf data/supplier_invoice_2026-05.pdf data/supplier_invoice_2026-06.pdf \
  --households data/households.csv \
  --xero-contacts data/Lot_Names_and_Addresses.csv
```

No spreadsheet is needed, and there is no copying and pasting between files.

---

## Contents

1. [What it does](#what-it-does)
2. [What's in this repository](#whats-in-this-repository)
3. [Installation](#installation)
4. [Quick start (try it with the example data)](#quick-start-try-it-with-the-example-data)
5. [Input files](#input-files)
6. [Choosing the billing months](#choosing-the-billing-months)
7. [All command-line options](#all-command-line-options)
8. [More examples](#more-examples)
9. [Output files](#output-files)
10. [How the calculation works](#how-the-calculation-works)
11. [Special situations](#special-situations)
12. [Troubleshooting](#troubleshooting)
13. [Keeping real data private](#keeping-real-data-private)

---

## What it does

A cluster is a group of households that share a battery (BESS), some common
loads (lighting, pumps, etc.) and **one** connection to the electricity supplier.
The supplier sends **one** invoice for the whole cluster each month, and that cost
has to be split fairly between the households. This script does the split.

```
 households.csv ─────────────┐
 meter export CSVs (monthly) ┼──►  cluster_energy_billing.py  ──►  one PDF invoice per household
 supplier invoice PDFs ──────┤                                 ──►  monthly calculation workbook (.xlsx)
 Lot_Names_and_Addresses.csv ┘                                 ──►  Xero invoice + credit-note CSVs
                                                               ──►  audit CSVs
```

Each run:

1. **Reads the meters.** The meter exports contain *running totals* (like a car
   odometer), so each month's usage is "this month's reading minus last month's".
2. **Reads the supplier PDFs.** It finds the *Usage Calculations* table on each
   invoice and extracts every charge line automatically.
3. **Splits the costs.** Energy each household exports is credited, part at the
   microgrid (peer-to-peer) rate and part at the grid feed-in rate. Energy it imports is
   charged, part at the microgrid rate and part at the grid rate. Shared supplier fees
   and common loads are divided between the households.
4. **Reconciles.** Any leftover difference is shared out so that the household
   invoices add up **exactly** to the supplier's bill, to the cent.
5. **Writes the outputs:** invoices, workbook, Xero files and audit trail.

You can bill one month or several (for example a quarter); multi-month invoices
show the total plus a month-by-month breakdown.

---

## What's in this repository

```
wev_billing/
├── cluster_energy_billing.py      ← the program (one file)
├── requirements.txt               ← Python packages it needs
├── README.md                      ← this guide
├── .gitignore                     ← stops invoices and real data being committed
└── data/                          ← FICTITIOUS example data to practise with
    ├── households.csv
    ├── Lot_Names_and_Addresses.csv
    ├── meters_2026-03.csv         ← March: only provides the opening readings
    ├── meters_2026-04.csv
    ├── meters_2026-05.csv
    ├── meters_2026-06.csv
    ├── supplier_invoice_2026-04.pdf
    ├── supplier_invoice_2026-05.pdf
    └── supplier_invoice_2026-06.pdf
```

All names, addresses, emails, account numbers and readings in `data/` are made up.

---

## Installation

### 1. Install Python (3.9 or newer)

Check whether you already have it. Open a terminal (on Windows, *Command Prompt* or
*PowerShell*) and type:

```bash
python3 --version
```

On Windows, type `py --version` instead. If you see `Python 3.9` or higher, you're set.
Otherwise:

| System | How to install Python |
|---|---|
| **macOS** | Download the installer from [python.org/downloads](https://www.python.org/downloads/), or with Homebrew: `brew install python` |
| **Windows** | Download from [python.org/downloads](https://www.python.org/downloads/). In the installer, tick **"Add python.exe to PATH"**. Then use `py` wherever this guide says `python3`. |
| **Linux (Ubuntu/Debian)** | `sudo apt install python3 python3-venv` |

### 2. Get the code

Either download the repository as a ZIP from GitHub (green **Code** button →
**Download ZIP**) and unzip it, or clone it:

```bash
git clone https://github.com/dbarneche/wev_billing.git
cd wev_billing
```

All commands in this guide are run from inside this folder.

### 3. Python packages: automatic

The script needs a few packages (`numpy`, `pandas`, `openpyxl`, `reportlab` and,
sometimes, `pdfplumber`). **You don't have to install them yourself.** On the first
run the script checks what's missing and installs it:

- **If you're inside a virtual environment**, it installs the packages there.
- **Otherwise**, it creates a private environment called `.billing-venv` next to
  the script, installs the packages into it, and re-runs itself inside it. This avoids
  the *"externally-managed-environment"* error that many newer Python installations
  show. Later runs reuse it, and it doesn't touch your system Python.

The first run needs an internet connection and may take a minute.

<details>
<summary><b>Prefer to install the packages yourself?</b></summary>

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
python3 -m pip install -r requirements.txt
```

Then run the script with `--no-auto-install` if you want it to only *report*
missing packages instead of installing them.
</details>

### 4. Optional: poppler's `pdftotext`

The supplier PDFs are read with `pdftotext` (part of *poppler*) if it's installed,
otherwise with the `pdfplumber` package. Both work; `pdftotext` is a little more
robust. To install it: `brew install poppler` (macOS),
`sudo apt install poppler-utils` (Linux), or skip it on Windows.

---

## Quick start (try it with the example data)

From inside the repository folder (macOS / Linux):

```bash
python3 cluster_energy_billing.py --months Apr-Jun \
  --meters data/meters_*.csv \
  --pdfs data/supplier_invoice_*.pdf \
  --households data/households.csv \
  --xero-contacts data/Lot_Names_and_Addresses.csv
```

> Use `data/meters_*.csv`, **not** `data/*.csv`: the folder also holds
> `households.csv` and `Lot_Names_and_Addresses.csv`, which are not meter files.

<details>
<summary><b>Windows (PowerShell) version</b></summary>

PowerShell doesn't expand `*.csv` for other programs, so list the files and use a
backtick `` ` `` to continue a line:

```powershell
py cluster_energy_billing.py --months Apr-Jun `
  --meters data/meters_2026-03.csv data/meters_2026-04.csv `
           data/meters_2026-05.csv data/meters_2026-06.csv `
  --pdfs data/supplier_invoice_2026-04.pdf `
         data/supplier_invoice_2026-05.pdf `
         data/supplier_invoice_2026-06.pdf `
  --households data/households.csv `
  --xero-contacts data/Lot_Names_and_Addresses.csv
```
</details>

You'll see a short report. The last line confirms that the household invoices add
up to the supplier bills:

```
Billing months: Apr 2026, May 2026, Jun 2026
Households: 24 lots from households.csv
Supplier PDF supplier_invoice_2026-04.pdf: EXAMPLE2B202604, 01/04/2026 to 01/05/2026, total $547.73 ...
Apr 2026: supplier $547.73 | households $547.73 | variance $117.10 | extra meters: DB1=6 kWh, DB2=14 kWh
...
Wrote 24 invoices to .../invoices
Period total: households $2,340.76 vs supplier $2,340.76.
```

Everything is written to a new `invoices/` folder (see [Output files](#output-files)).

> **Why four meter files for three months?** See the next section. It's the most
> important rule in this guide.

---

## Input files

### The golden rule for meter files

Meters report **running totals**. To know how much was used in April, the script
needs the reading at the **end of March** as well as the end of April.

> **Always include the meter CSV for the month _before_ your first billing month.**
> To bill April–June, pass the March, April, May and June files.

If you forget, the script stops and tells you which file to add. (An alternative
to the extra CSV is the `meter_readings_ledger.csv` file from a previous run; see
[`--opening-readings`](#run-inputs).)

### 1. `households.csv`: one row per household

Pass its location with `--households`, e.g. `--households data/households.csv`.
(Without the option, the script looks for `households.csv` **next to the script**,
not in `data/`, and stops with *"Households file not found"* if it isn't there.)

| Column | Required | Meaning | Example |
|---|---|---|---|
| `Cluster` | ✅ | Cluster code. Use digits optionally followed by one letter (e.g. `2B`, `1A`, `7`) so it matches the meter names. | `2B` |
| `Lot` | ✅ | Lot name, `Lot` followed by a number | `Lot 01` |
| `Owners` | ✅ | Name printed on the invoice | `Alex Archer & Jamie Archer` |
| `Unit Entitlement` | ✅ | The lot's unit entitlement (UE). Shares are UE ÷ total UE. | `36` |
| `Energy Meter Installed` | ✅ | `Y` or `N`. Only `Y` households pay the supplier's pass-through fees. | `Y` |
| `Street Address`, `Suburb`, `State`, `Post Code` | | Printed on the invoice | `1 Banksia Way`, `Exampleville`, `WA`, `6999` |
| `Block Type` | | Printed on the invoice | `Cottage` |
| `Export Capacity (kW)` | | Solar export capacity, shown for information | `5.0` |
| `PV Installed`, `Functional Droplet` | | `Y`/`N`, shown for information | `Y` |

### 2. `Lot_Names_and_Addresses.csv`: Xero contact names

This is simply the **contacts export from Xero** (Contacts → Export). Only two
columns are used:

| Column | Meaning | Example |
|---|---|---|
| `*ContactName` | The contact's name **exactly as it appears in Xero** | `Alex Archer & Jamie Archer` |
| `AccountNumber` | `LOT` + lot number | `LOT1` |

Pass its location with `--xero-contacts`, e.g. `--xero-contacts data/Lot_Names_and_Addresses.csv`.
(Without the option, the script looks **next to the script**. If the file isn't
found, the household `Owners` names are used as Xero contact names and a note is
printed. Make sure they match your Xero contacts before importing.)

### 3. Meter export CSVs: one file per month

These are the monthly exports from the metering platform (Stormcloud). Each file
holds the **reading at the end of the month** for every meter.

| Column | Meaning |
|---|---|
| `unit_name` | Meter name (see the naming rules below) |
| `Period Start` | Start of the month, e.g. `2026-04-01T08:00:00.000Z`. **This decides which month the file belongs to.** |
| `Period End` | End of the month (for information only) |
| `Meter Import (kWh)` | Running total of energy taken **from** the microgrid |
| `Meter Export (kWh)` | Running total of energy sent **to** the microgrid (solar) |

The column order doesn't matter, and extra empty columns are ignored.

**Meter naming rules.** The script recognises each meter from its name:

| Name contains… | Treated as | Example |
|---|---|---|
| `Lot` + number (after the cluster code) | That household | `2B Lot 01 - Alex Archer` |
| `BESS` or `Battery` | The shared battery | `2B SMSB BESS Meter` |
| `Common Serv…` | Common services (shared loads) | `2B SMSB Common Services` |
| `Grid` (or just `SMSB`) | The main grid connection meter (used for checks only) | `2B SMSB Grid Meter` |
| `EV` | Shared EV charger | `2B EV Charger` |
| anything else | An **extra microgrid meter**, shared like common services | `2B DB1`, `2B DB2` |

> The metering platform reports the **battery's import and export columns the wrong
> way round**. The script corrects this automatically (`--swap-import-export BESS`).

### 4. Supplier invoice PDFs: one per billing month

The monthly invoices from your electricity supplier, as **text-based PDFs**
(downloaded, not scanned). The script reads:

- the **Usage Calculations** table: each charge line with its usage, rate and amount,
  down to `Total This NMI`;
- the **supply period** (`Start Time:` / `Supply period:`), which decides the billing month;
- the **account reference number**.

Charge lines are recognised by name: *On/Off Peak Energy*, *Daily Supply*,
*Essential System Services*, *Solar On/Off Peak Feedback*, *STP/RTP*,
*Western Power / Network Tariff*, *Market Fee*. Any other line is counted as a
supplier fee, and a note is printed so you can check it.

If an invoice contains more than one *Usage Calculations* block, the first one is
used; pick a specific meter point with `--nmi`.

---

## Choosing the billing months

`--months` accepts several formats:

| You type | Months billed |
|---|---|
| `--months Jun` | June |
| `--months Apr-Jun` | April, May, June |
| `--months "Apr,May,Jun"` | April, May, June |
| `--months 4-6` | April, May, June |
| `--months "Jan:Dec"` | the whole year |

The year comes from `--year` (default `2026`). Remember to pass **one supplier PDF
per billing month** and **the meter CSVs for those months plus the month before**.

---

## All command-line options

Only `--months`, `--meters` and `--pdfs` are required; everything else has a
sensible default. Run `python3 cluster_energy_billing.py --help` for the same list
in your terminal.

### Global parameters (tariffs and cluster settings)

| Option | Default | Meaning |
|---|---|---|
| `--grid-export-rate` | `0.0385` | $/kWh (inc. GST) credited for energy exported to the main grid |
| `--grid-import-rate` | `0.2994` | $/kWh (inc. GST) charged for energy imported from the main grid |
| `--p2p-rate` | `0.20` | $/kWh for energy traded inside the microgrid (peer-to-peer) |
| `--num-households` | `24` | Number of households sharing the supplier's fixed costs |
| `--no-pv-p2p-rate` | *(blank)* | Accepted for compatibility with the original spreadsheet; not used in any calculation |
| `--battery-capacity` | `232` | Shared battery capacity in kWh (sets each household's battery allocation) |
| `--export-limit` | `0.66568` | Export limit factor (limited export = export capacity × factor; shown on invoices) |
| `--gst` | `0.10` | GST rate applied to supplier charges (0.10 = 10%) |

### Run inputs

| Option | Default | Meaning |
|---|---|---|
| `--year` | `2026` | Calendar year of the billing months |
| `--months` | **required** | Months to bill, e.g. `Apr-Jun` (see [Choosing the billing months](#choosing-the-billing-months)) |
| `--meters` | **required** | Meter export CSVs: one per billing month **plus the month before** (up to 13 files) |
| `--pdfs` | **required** | Supplier invoice PDFs, one per billing month |
| `--opening-readings` | *(none)* | Instead of the extra month's CSV: the `meter_readings_ledger.csv` from a previous run |
| `--households` | `households.csv` next to the script | Path to the households table, e.g. `data/households.csv` |

### Behaviour

| Option | Default | Meaning |
|---|---|---|
| `--extra-assets` | `common` | What to do with extra microgrid meters (e.g. `DB1`, `DB2`). `common` = share their energy by unit entitlement, like common services. `ignore` = show them in reports but charge nothing (gives the same numbers as the original spreadsheet). |
| `--new-asset-baseline` | `zero` | For an extra or EV meter with **no previous reading** (e.g. newly installed). `zero` = assume it started at 0 kWh. `skip` = leave it out that month. |
| `--swap-import-export` | `BESS` | Comma-separated meters whose import/export columns are reversed in the CSVs. Use `none` to turn this off. |
| `--timezone` | `Australia/Perth` | Local timezone used to decide which month each CSV belongs to |
| `--nmi` | *(first block)* | Supplier meter number (NMI) to read, if a PDF contains several |

### Output

| Option | Default | Meaning |
|---|---|---|
| `--outdir` | `invoices` | Folder for all outputs (created if missing) |
| `--issue-date` | today | Invoice issue date, `YYYY-MM-DD` |
| `--due-days` | `14` | Days from the issue date until payment is due |
| `--title` | `Cluster Energy Billing` | Heading printed at the top of every invoice |
| `--issuer` | `The Owners of the Cluster` | Issuer line printed under the heading |
| `--no-auto-install` | off | Don't install missing Python packages; just list them and stop |

### Xero export

| Option | Default | Meaning |
|---|---|---|
| `--no-xero` | off | Don't write the Xero CSV files |
| `--xero-contacts` | `Lot_Names_and_Addresses.csv` next to the script | Xero contacts export (matches each lot to its Xero contact), e.g. `data/Lot_Names_and_Addresses.csv` |
| `--xero-invoice-suffix` | `/MMYY` of the invoice date | Ending of the Xero invoice numbers, e.g. `/0326` gives `2B6/0326`. **Must be different for every billing run**, because Xero rejects duplicate numbers. |
| `--xero-invoice-date` | `--issue-date` | Invoice date written to Xero (`YYYY-MM-DD`) |
| `--xero-due-date` | invoice due date | Due date written to Xero (`YYYY-MM-DD`) |
| `--xero-account-charges` | `409` | Xero account code for "Charges and Fees" |
| `--xero-account-imports` | `410` | Xero account code for "Total Imports" |
| `--xero-account-exports` | `411` | Xero account code for "Total Exports" |
| `--xero-tax-type` | `BAS Excluded` | Xero tax type for every line |
| `--xero-branding-theme` | `Energy Invoice` | Xero branding theme for the invoices |

---

## More examples

> macOS/Linux shown. On Windows use `py` instead of `python3`, list files instead of
> using `*` or `{…}`, and end continued lines with `` ` `` instead of `\`.
>
> The April–June examples run as-is on the example data in `data/`. The July
> examples show the pattern for your own files (the repository has no July data).
> Every example passes `--households` and `--xero-contacts`, because those files
> live in `data/`, not next to the script.

**Bill a single month** (June needs the May file for its opening readings):

```bash
python3 cluster_energy_billing.py --months Jun \
  --meters data/meters_2026-05.csv data/meters_2026-06.csv \
  --pdfs data/supplier_invoice_2026-06.pdf \
  --households data/households.csv --xero-contacts data/Lot_Names_and_Addresses.csv
```

**Bill a quarter with your own heading, dates and Xero numbering:**

```bash
python3 cluster_energy_billing.py --year 2026 --months Apr-Jun \
  --meters data/meters_2026-0{3,4,5,6}.csv \
  --pdfs data/supplier_invoice_2026-0{4,5,6}.pdf \
  --households data/households.csv --xero-contacts data/Lot_Names_and_Addresses.csv \
  --title "Example Eco Village - Cluster 2B" \
  --issuer "The Owners of Cluster 2B, Strata Plan 00000" \
  --issue-date 2026-07-15 --due-days 21 \
  --xero-invoice-suffix /Q2-26
```

**Reproduce the original spreadsheet's numbers** (extra meters such as DB1/DB2 not charged):

```bash
python3 cluster_energy_billing.py --months Apr-Jun --extra-assets ignore \
  --meters data/meters_2026-0{3,4,5,6}.csv --pdfs data/supplier_invoice_2026-0{4,5,6}.pdf \
  --households data/households.csv --xero-contacts data/Lot_Names_and_Addresses.csv
```

**Invoices only, no Xero files, to a different folder:**

```bash
python3 cluster_energy_billing.py --months Apr-Jun --no-xero --outdir draft_q2 \
  --meters data/meters_2026-0{3,4,5,6}.csv --pdfs data/supplier_invoice_2026-0{4,5,6}.pdf \
  --households data/households.csv
```

**Change a tariff**, e.g. a new peer-to-peer rate (your own July files):

```bash
python3 cluster_energy_billing.py --months Jul --p2p-rate 0.22 \
  --meters private/meters_2026-06.csv private/meters_2026-07.csv \
  --pdfs private/supplier_invoice_2026-07.pdf \
  --households private/households.csv --xero-contacts private/Lot_Names_and_Addresses.csv
```

**Continue from last time's readings** instead of passing the previous month's CSV
(after a June run that wrote `invoices/meter_readings_ledger.csv`):

```bash
python3 cluster_energy_billing.py --months Jul \
  --opening-readings invoices/meter_readings_ledger.csv \
  --meters private/meters_2026-07.csv --pdfs private/supplier_invoice_2026-07.pdf \
  --households private/households.csv --xero-contacts private/Lot_Names_and_Addresses.csv \
  --outdir invoices_july
```

---

## Output files

Everything goes into the `--outdir` folder (default `invoices/`). File names
include the year and period, e.g. `2026_Apr-Jun`.

| File | What it is |
|---|---|
| `<cluster>_Lot<NN>_<year>_<period>.pdf` | **The invoice for each household** (A4, usually 2 pages): amount due or credit, energy summary, itemised charges with quantities and unit prices, month-by-month breakdown, how the share was calculated, cluster context and method notes. Amounts owed are shown in red, credits in green. |
| `monthly_tabs_<year>_<period>.xlsx` | Calculation workbook: one sheet per month (supplier charges, cluster figures, every household's full calculation), a combined *Master - Households Data* sheet, and a *Notes* sheet. Columns named "Amanda …" are the supplier's charges; the name comes from the original spreadsheet. |
| `invoices_xero_data.csv` | Ready to import into Xero (Business → Invoices → Import). |
| `credit_note_xero_data.csv` | Xero credit notes for households whose total is a credit. |
| `invoice_totals_<year>_<period>.csv` | One line per household with its total. |
| `cluster_monthly_summary_<year>_<period>.csv` | Cluster-level figures per month (supplier totals, battery, common loads, reconciliation). |
| `supplier_bill_lines_<year>_<period>.csv` | Every charge line read from the supplier PDFs. **Worth a quick check** that the PDFs were read correctly. |
| `meter_readings_ledger.csv` | All meter readings used (battery already corrected). Can be passed to `--opening-readings` next time. |
| `run_notes.txt` | Any warnings from the run (only written if there were some). |

### How the Xero files are built

Each household gets three lines for the billing period:

| Line | Contains | Account |
|---|---|---|
| Charges and Fees | daily supply share, supplier fees, common loads, reconciliation | 409 |
| Total Imports | microgrid + main grid import charges | 410 |
| Total Exports | microgrid + main grid export credits (negative) | 411 |

If the household **owes money**, all three lines go on one Xero invoice. If its
total is a **credit**, the first two lines form an invoice and the exports line
becomes a credit note (same number + `CN`). Rounding is handled so that every
Xero invoice equals its PDF invoice to the cent.

---

## How the calculation works

Everything below is done separately for each month, then added up over the billing
period. "UE share" means a household's unit entitlement ÷ total unit entitlement.

**1. Monthly energy per meter**

```
usage this month = closing reading − previous month's closing reading
```

**2. Supplier figures** (from the PDF)

```
supplier import       = on-peak kWh + off-peak kWh
supplier solar credit = solar on-peak + solar off-peak feedback kWh
fixed fees            = (ESS + STP + RTP + network + market fee + other) × (1 + GST)
daily supply          = daily supply charge × (1 + GST)
```

**3. Exports: who supplied the microgrid?**

Some of the solar the households export is used inside the cluster (the
*microgrid export pool*). The rest goes out to the grid.

```
microgrid export pool  = total household export − supplier solar credit
battery allocation     = battery capacity × UE share                 (kWh per day)
applied allocation     = min(household export, battery allocation × supply days)
microgrid contribution = pool × applied allocation ÷ total applied allocation

microgrid export = microgrid contribution          → credited at --p2p-rate
grid export      = household export − microgrid export → credited at --grid-export-rate
```

**4. Imports: where did each household's energy come from?**

```
BESS efficiency  = battery export ÷ battery import       (round-trip efficiency)
microgrid import = min(microgrid contribution × BESS efficiency, household import)
grid import      = household import − microgrid import

microgrid import → charged at --p2p-rate
grid import      → charged at --grid-import-rate
```

**5. Shared costs**

```
daily supply share = daily supply ÷ --num-households
fixed fees share   = fixed fees ÷ --num-households      (only if Energy Meter Installed = Y)
common loads share = common load kWh × UE share × rate
    rate = --p2p-rate if the cluster exported more to the grid than it imported,
           otherwise --grid-import-rate
    (common loads = common services, plus DB1, DB2 etc. with --extra-assets common)
EV charger share   = EV kWh × UE share × --p2p-rate
```

**6. Reconciliation: making it add up exactly**

```
subtotal       = all charges − all credits
variance       = supplier invoice total − sum of all households' subtotals
variance share = variance × (household import ÷ total household import)
final invoice  = subtotal + variance share
```

So the household invoices always add up **exactly** to the supplier's invoice.

---

## Special situations

### Extra microgrid meters (DB1, DB2, …)

Any meter that isn't a lot, the battery, common services, the grid meter or the EV
charger is treated as an **extra microgrid meter**. By default its energy is shared
by UE like common services and appears as its own line on the invoices. Use
`--extra-assets ignore` to report it without charging it.

### A meter with no previous reading (newly installed)

For a battery, common-services or **household** meter this is an error: add the
previous month's CSV. For an extra or EV meter:

- `--new-asset-baseline zero` (default) assumes its total started at 0 kWh. This is
  correct if it was installed during the month. If it had already been running, it
  would bill everything it has recorded so far, so check the note printed by the script.
- `--new-asset-baseline skip` leaves it out until a previous reading exists.

### A household meter is replaced or goes offline

The script needs an unbroken running total for every household. If a meter goes
dark or is replaced:

1. **Estimate the missing usage.** A defensible method is to assume the household
   keeps its recent share of the other households' usage:
   `estimate = (household kWh ÷ other households' kWh, over the last 3 complete months)
   × other households' kWh in the missing month`.
2. **Write a continued running total** in the CSV: last reliable reading + estimate.
3. After a **replacement** (new meter starting at 0), add a fixed offset to every
   new reading so the total keeps going up:
   `offset = continued running total at the installation date`.

### Checking against the original spreadsheet

If you're comparing with the original spreadsheet, run with
`--extra-assets ignore`. The spreadsheet ignores DB1/DB2, and with this flag the
figures match it in every column. If they still differ, look for a broken formula
in the spreadsheet's meter-readings tab. One wrong cell reference changes every
household's result, because the shares are calculated across the whole cluster.

### Xero invoice numbers

Numbers look like `2B6/0326` (cluster `2B`, lot `6`, suffix `/0326`). Choose a new
`--xero-invoice-suffix` for every billing run so numbers never repeat.

---

## Troubleshooting

| Message | What it means | Fix |
|---|---|---|
| `No opening read for '…' in Apr 2026` | The previous month's reading is missing. | Add the previous month's meter CSV to `--meters` (or a ledger with `--opening-readings`). |
| `…: no column containing 'period start'` | A file passed to `--meters` isn't a meter export (e.g. `data/*.csv` also picked up `households.csv`). | Pass only the meter files, e.g. `--meters data/meters_*.csv`. |
| `Households file not found` | No `--households` option was given, so the script looked next to itself. | Add `--households data/households.csv` (or wherever your file is). |
| `No supplier PDF for: Jul 2026` | A billing month has no supplier invoice. | Add that month's PDF to `--pdfs`, or remove the month from `--months`. |
| `More than one supplier PDF for …` | Two PDFs cover the same month. | Pass only one per month. |
| `no 'Usage Calculations' section found` | The PDF isn't a supplier invoice in the expected format, or it's a scanned image. | Download the original PDF from the supplier. |
| `NOTE: … unrecognised supplier line(s)` | A charge line had an unfamiliar name. | Usually fine: it's counted as a supplier fee. Check `supplier_bill_lines_*.csv`. |
| `NOTE: … non-numeric reads ignored for …` | A meter had a blank or invalid reading in the CSV. | Fill it in. For a household, the run stops until it's provided. |
| `NOTE: … register went backwards` | A running total decreased (meter reset or replaced). | See [A household meter is replaced](#a-household-meter-is-replaced-or-goes-offline). |
| `Households table has N lots but --num-households is 24` | The two numbers differ. | Set `--num-households` to your number of households. |
| `could not install … automatically` | No internet access, or pip is blocked. | Connect to the internet, or install manually (see [Installation](#installation)). |
| `python3: command not found` (Windows) | Windows uses a different command. | Use `py` instead of `python3`. |
| `externally-managed-environment` | Your system Python refuses direct installs. | Nothing to do: the script creates its own `.billing-venv` automatically. |

---

## Keeping real data private

Meter readings, invoices and contact lists contain personal information, so keep
your **real** files out of Git. The included `.gitignore` already ignores:

- generated outputs: `invoices/` and any folder starting with `out`;
- the private `.billing-venv/` and `.venv/` environments;
- any `.csv`, `.pdf` or `.xlsx` file in the **top** folder;
- a `private/` folder, for your real data.

The fictitious example files in `data/` **are** tracked, so don't put real data there.

A simple way to work:

```
wev_billing/
├── cluster_energy_billing.py
├── data/                          ← fictitious examples (tracked by Git)
└── private/                       ← your real files (ignored by Git)
    ├── households.csv
    ├── Lot_Names_and_Addresses.csv
    ├── meters_2026-06.csv
    ├── meters_2026-07.csv
    └── supplier_invoice_2026-07.pdf
```

Then point the options at it, e.g. `--households private/households.csv`.

Before pushing, run `git status` and make sure no real data is listed.
