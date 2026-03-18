# CLAUDE.md — Kofeed Proficiency Test Platform

## Project Overview

This is a **Korean feed industry proficiency testing (숙련도 시험) platform** built with Streamlit. Member companies submit analytical data for feed samples; the system computes Robust Z-scores for inter-laboratory comparison and distributes PDF reports via email.

The app is deployed on **Streamlit Community Cloud** and uses **Google Sheets** as both database and configuration store.

---

## Repository Structure

```
kofeed/
├── app.py                  # Main participant data-submission form (entry point)
├── pages/
│   └── admin.py            # Admin dashboard (password-protected)
├── utils/
│   ├── __init__.py
│   ├── config.py           # All config read/write; Google Sheets "config" tab
│   ├── sheets.py           # Submission data read/write; Google Sheets "제출데이터" tab
│   ├── zscore.py           # Robust Z-score computation helpers
│   ├── report.py           # PDF generation (ReportLab, Korean font)
│   └── email_sender.py     # SMTP email dispatch (Gmail)
├── requirements.txt        # Python dependencies
├── packages.txt            # System packages (fonts-nanum for Korean PDF)
└── .gitignore
```

---

## Tech Stack

| Layer | Technology |
|---|---|
| UI framework | Streamlit ≥ 1.35 |
| Data storage | Google Sheets (via gspread ≥ 6.0) |
| Auth | Google service account (`st.secrets["gcp_service_account"]`) |
| PDF generation | ReportLab ≥ 4.2 |
| Korean font | NanumGothic TTF (`packages.txt`: `fonts-nanum`) |
| Email | SMTP over Gmail SSL port 465 |
| Data processing | pandas ≥ 2.0, numpy ≥ 1.26 |

---

## Secrets Configuration

Secrets live in `.streamlit/secrets.toml` (gitignored). Required keys:

```toml
[gcp_service_account]
# Full Google service account JSON contents as TOML keys

[sheet]
name = "스프레드시트 파일명"   # Google Sheets file name

[admin]
password = "admin_password"

[email]
sender   = "sender@gmail.com"
password = "gmail_app_password"
```

---

## Google Sheets Layout

Two worksheets are used:

### `config` tab
All application configuration is stored here. Columns: `type | group | name | samples | order | enabled | use_equip | use_solvent | free_decimal`

Config row types:

| `type` | Purpose | Key fields |
|---|---|---|
| `method_option` | Analysis method dropdown items | `name`=method label, `group`=component name (empty=global) |
| `info_field` | Participant info form fields | `name`=field label, `group`=placeholder, `samples`=flags (`required`, `email`) |
| `sample` | Feed sample types (사료 종류) | `name`=sample name |
| `group` | Component sections | `name`=section name, `samples`=`nir` if NIR-eligible |
| `component` | Individual analytes | `group`=section, `name`=analyte, `samples`=`all` or comma-separated sample names |
| `question` | Optional survey questions | `group`=question ID, `name`=question text, `samples`=type spec (see below) |
| `participant` | Access code → company name map | `group`=code, `name`=company name |

Question `samples` type specs:
- `text` or `text:hint` → free-text
- `choice:opt1|opt2|opt3` → single radio
- `multicheck:opt1|opt2|opt3` → multi-checkbox

### `제출데이터` tab
Submission rows. Header row is auto-created on first submit; new columns are appended automatically when a new component is added to config. Column naming conventions:
- `{component}_{sample}` — numeric value (e.g., `수분_축우사료`)
- `{component}_방법` — analysis method string
- `{component}_기기` — instrument name
- `{component}_용매` — solvent name
- `NIR_{component}_{sample}` — NIR measurement value

---

## Application Flow

### Participant Form (`app.py`)
1. If `participant` entries exist in config, show a participant-code gate before the form.
2. Load config from Google Sheets (cached 120s via `@st.cache_data(ttl=120)`).
3. Render institution info fields, component input tables per group/section, optional survey questions.
4. On submit: validate required fields and method-selection rules, then call `submit_data()` to append to Google Sheets, then send confirmation email with PDF attachment.

Component input table columns: **성분 | 방법 (dropdown) | 기기명 | 용매 | [sample columns...]**

### Admin Dashboard (`pages/admin.py`)
Protected by `st.secrets["admin"]["password"]`. Four tabs:
- **제출 현황** — view all submissions as a dataframe
- **Z-score 분석** — Robust Z-score analysis per analyte (overall and by method), with color-coded tables
- **보고서 발송** — generate individual PDF reports (overall + by-method), download buttons, and bulk email dispatch
- **설정** — full CRUD for all config types via `st.data_editor`

---

## Robust Z-Score (`utils/zscore.py`)

Formula: `Z = (x − median) / (1.4826 × MAD)`

- Falls back to standard-deviation Z-score when MAD = 0.
- **Minimum n > 5** required to compute Z-scores (returns `NaN` otherwise).
- By-method Z-scores also require > 5 institutions using the same method.
- Thresholds: `|Z| ≤ 2.0` = 적합 (pass), `2.0 < |Z| ≤ 3.0` = 경고 (warning), `|Z| > 3.0` = 부적합 (fail).

---

## PDF Reports (`utils/report.py`)

Three report types:
- **전체 Z-score** (`generate_pdf_overall`) — compares against all institutions.
- **방법별 Z-score** (`generate_pdf_by_method`) — compares within same analysis method group.
- **제출 확인서** (`generate_submission_pdf`) — sent immediately on submission; no Z-scores.

Korean font: loads `/usr/share/fonts/truetype/nanum/NanumGothic.ttf`. Falls back to Helvetica if not found (local dev without `fonts-nanum` installed).

---

## Key Conventions

### Session State Keys
Component input widgets use the key pattern `{group_name}_{comp}_{sample}` (e.g., `일반성분_수분_축우사료`). On form submission, values are read directly from `st.session_state` using these keys to avoid stale data.

### Column Naming
- `is_value_col(col, samples)` — returns True if `col` ends with `_{sample_name}`.
- `get_component_from_col(col, samples)` — strips sample suffix (and `NIR_` prefix if present).
- `get_sample_from_col(col, samples)` — returns the sample portion of the column name.

### Institution Field Detection
The institution/company name field is found by keyword matching: any `info_field` whose `name` contains `기관`, `회사`, `기업`, or `업체`. The first such field wins; defaults to the first info field.

### Cache Invalidation
- `get_config()` has `ttl=120` seconds.
- `load_data()` in admin has `ttl=60` seconds.
- Call `st.cache_data.clear()` or `get_config.clear()` to force refresh (done automatically after `save_config()`).

### Boolean Column Handling
`use_equip`, `use_solvent`, and `free_decimal` columns default to `True`/`False` and tolerate missing columns in older sheets — new columns added gracefully.

---

## Development Workflow

### Running Locally

```bash
pip install -r requirements.txt
# Install system font (Linux):
sudo apt-get install fonts-nanum

# Create .streamlit/secrets.toml with required keys (see above)
streamlit run app.py
```

### Adding a New Analyte/Component
1. In the admin **설정** tab, add a row to **성분 종목** with the appropriate `group`, `name`, `samples`, and optional flags.
2. Click **설정 저장** — the config sheet is updated and the form refreshes within 2 minutes (cache TTL).
3. New submission columns are appended to the data sheet automatically on the next submission.

### Adding a New Sample (Feed Type)
1. Add a row to **사료 종류** in admin settings.
2. Update component rows' `samples` field if the new sample type is not covered by `all`.

### Modifying PDF Layout
Edit `utils/report.py`. The `_make_table_style()` function centralizes table styling. Column widths are in `mm` units.

### Email Configuration
Email uses Gmail SMTP with an App Password (not the account password). Set `email.sender` and `email.password` in secrets. Korean filenames in attachments are RFC 2047 base64-encoded for compatibility with Naver/Daum mail.

---

## Deployment (Streamlit Community Cloud)

- Set all secrets in the Streamlit Cloud secrets manager (not committed to git).
- `packages.txt` installs system-level `fonts-nanum` for Korean PDF rendering.
- The app auto-restarts on each push to the connected branch.
- The Google service account must have **Editor** access to the target spreadsheet.
