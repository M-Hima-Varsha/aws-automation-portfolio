# AWS Automation Portfolio

## A collection of AWS reporting and billing automation projects

| Field      | Details                    |
|------------|----------------------------|
| Owner      | Hima Varsha M              |
| Repository | `aws-automation-portfolio` |
| Branch     | `cost-automation-codes`    |

---

## 1. Overview

This branch contains AWS automation projects that eliminate manual effort in billing and reporting tasks.
Each automation reads input files, processes the data, and generates a structured Excel report as output.

### Automations available in this branch:

| #  | Automation Name       | What it does                                              |
|----|-----------------------|-----------------------------------------------------------|
| 1  | Consolidated Report   | Merges monthly AWS billing CSVs into one Excel report     |
| 2  | Invoice Automation    | Extracts data from AWS invoice PDFs into an Excel report  |
| 3  | RI Utilization Report | Summarizes Reserved Instance utilization into Excel       |

---

## 2. Branch Structure

> You are currently on the `cost-automation-codes` branch.
> All folders and scripts listed below are part of this branch.

```text
cost-automation-codes/
|
|-- README.md                          <- This file. Branch-level documentation
|-- main.gitignore                     <- Git ignore rules for the entire branch
|
|-- Consolidated Report/               <- Automation 1
|   |-- README.md                      <- Documentation for this automation
|   |-- consolidated-automation.py     <- Main script to run
|   |-- accounts.txt                   <- List of AWS account names/IDs
|   |-- Month1/                        <- Input folder for Month 1 CSV files
|   |   `-- 123456789012_account_report.csv
|   `-- Month2/                        <- Input folder for Month 2 CSV files
|       `-- 123456789012_account_report.csv
|
|-- Invoice Automation/                <- Automation 2
|   |-- README.md                      <- Documentation for this automation
|   |-- invoice-automation.py          <- Main script to run
|   |-- .gitignore                     <- Ignores invoice PDFs and Excel outputs
|   |-- assets/
|   |   `-- screenshots/               <- Images used in README documentation
|   |       `-- .gitkeep
|   `-- invoices/                      <- Place input invoice PDFs here
|       `-- .gitkeep
|
`-- RI Utilization Report/             <- Automation 3
    |-- README.md                      <- Documentation for this automation
    |-- utilization-automation.py      <- Main script to run
    `-- Month - YYYY/                  <- Input folder for RI utilization CSVs
        `-- reservations-utilization-table - *.csv
```

---

## 3. Automation Modules

### 3.1 Consolidated Report

**Purpose:** Reads monthly AWS billing CSV files for multiple accounts and combines them into a single consolidated Excel report.

| Field  | Details                                              |
|--------|------------------------------------------------------|
| Folder | `Consolidated Report/`                               |
| Script | `consolidated-automation.py`                         |
| Input  | Monthly CSV files placed in `Month1/`, `Month2/`     |
| Config | `accounts.txt` — contains the list of account names  |
| Output | Consolidated Excel report                            |

**How it works:**
1. Place monthly billing CSV files inside the respective month folders (`Month1/`, `Month2/`)
2. Add account names or IDs in `accounts.txt`
3. Run `consolidated-automation.py`
4. Excel report is generated as output

---

### 3.2 Invoice Automation

**Purpose:** Reads AWS invoice PDF files, extracts account-wise billing details, and generates a structured Excel Invoice Report.

| Field  | Details                                                              |
|--------|----------------------------------------------------------------------|
| Folder | `Invoice Automation/`                                                |
| Script | `invoice-automation.py`                                              |
| Input  | Invoice PDF files placed inside `invoices/`                          |
| Output | Excel Invoice Report                                                 |
| Extras | `assets/screenshots/` — stores images used in README documentation   |

**How it works:**
1. Place AWS invoice PDF files inside the `invoices/` folder
2. Run `invoice-automation.py`
3. Excel report is generated as output


---

### 3.3 RI Utilization Report

**Purpose:** Processes Reserved Instance utilization CSV exports from AWS Cost Explorer and generates a summarized Excel report.

| Field  | Details                                                        |
|--------|----------------------------------------------------------------|
| Folder | `RI Utilization Report/`                                       |
| Script | `utilization-automation.py`                                    |
| Input  | RI utilization CSV files placed in `Month - YYYY/`             |
| Output | RI Utilization Excel report                                    |

**How it works:**
1. Export RI utilization reports from AWS Cost Explorer
2. Place the CSV files inside the `Month - YYYY/` folder
3. Run `utilization-automation.py`
4. Excel report is generated as output

---

## 4. GitHub Upload Guidelines

Only source code, documentation, and placeholder files should be committed to this branch.

**Commit these files:**

```text
README.md
main.gitignore
automation-folder/README.md
automation-folder/*.py
automation-folder/input/.gitkeep
automation-folder/output/.gitkeep
```

**Do not commit confidential files:**

```text
*.pdf       <- Real invoice PDFs
*.xlsx      <- Generated Excel reports
*.csv       <- Real billing data
*.json      <- Config or credential files
*.env       <- Environment variable files
```

> Real AWS billing reports, invoice PDFs, generated Excel files, and credential files must stay local and should never be pushed to GitHub.

---

## 5. Documentation Standard for Each Automation

Every automation folder has its own `README.md` that covers:

- Purpose of the automation
- Access required
- Input files required
- Setup instructions
- Run command
- Output file details
- Common issues and fixes
- Owner details

---

## 6. Ownership

This branch and its source code are owned by **Hima Varsha M**.
