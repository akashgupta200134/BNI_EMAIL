# 📧 Outlook Web Email Automation (Playwright)

Automates sending personalised bulk emails via **Outlook Web** using Playwright.
Reads contacts from Excel Sheet1, templates from Sheet2, tracks progress,
and respects daily sending limits.

---

## 📁 Folder Structure

```
outlook_emailer/
├── main.py                  ← Entry point — run this
├── config.py                ← All settings (delays, limits, selectors)
├── requirements.txt
│
├── core/
│   └── email_sender.py      ← Playwright logic: compose, send, retry
│
├── utils/
│   ├── excel_handler.py     ← Read/write Excel (Sheet1 & Sheet2)
│   ├── state_manager.py     ← Daily send counter (state.json)
│   ├── validator.py         ← Email format validation
│   └── logger.py            ← File + console logging
│
├── scripts/
│   ├── generate_excel.py    ← Creates a fresh email_campaign.xlsx
│   └── report.py            ← Prints campaign progress summary
│
├── data/
│   ├── email_campaign.xlsx  ← Your Excel file (add contacts here)
│   └── state.json           ← Auto-created, tracks daily count
│
├── logs/
│   └── campaign_YYYY-MM-DD.log   ← Auto-created per day
│
└── screenshots/             ← Auto-saved on failures for debugging
```

---

## ⚙️ Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
playwright install chromium
```

### 2. Generate the Excel file

```bash
python scripts/generate_excel.py
```

This creates `data/email_campaign.xlsx` with sample data.
Open it and **fill in your real contacts in Sheet1**.

**Sheet1 columns:**
| Column | Description |
|--------|-------------|
| Name | Recipient's name (used as `{{Name}}` in template) |
| Email | Recipient's email address |
| Email Status | Auto-filled: `Invalid Email` or `Not Found` on failure |
| Remark | Auto-filled: `Done` after successful send. Pre-fill `Done` to skip. |

**Sheet2 columns:**
| Column | Description |
|--------|-------------|
| Template No | Label (Template 1, 2, 3…) |
| Subject Line | Email subject |
| Email Template Body | Full email body. Use `{{Name}}` for personalisation. |

---

## 🚀 Usage

### Run the campaign

```bash
python main.py
```

1. Browser opens Outlook Web
2. **You log in manually** (one time per session)
3. Press ENTER in the terminal
4. Emails are sent one-by-one with a 30-second gap

### Dry run (test without sending)

```bash
python main.py --dry-run
```

Opens the compose window and fills it, but does **not** click Send.

### Use a different Excel file

```bash
python main.py --excel /path/to/your_file.xlsx
```

### Reset daily counter

```bash
python main.py --reset-state
```

### Check campaign progress

```bash
python scripts/report.py
```

---

## ⚙️ Configuration (`config.py`)

| Setting | Default | Description |
|---------|---------|-------------|
| `DELAY_BETWEEN_EMAILS` | `30` | Seconds between each send |
| `BATCH_SIZE` | `50` | Emails per batch |
| `BATCH_PAUSE` | `60` | Seconds pause between batches |
| `DAILY_LIMIT` | `200` | Max emails per calendar day |
| `COMPOSE_RETRIES` | `2` | Retry attempts if compose fails |
| `HEADLESS` | `False` | Set `True` to hide browser window |

---

## 🔁 Resume After Crash

If the script crashes mid-run:
- Rows already marked `Done` will be **skipped** on the next run
- Daily counter in `state.json` is preserved
- Just run `python main.py` again — it picks up where it left off

---

## 🐛 Debugging Failures

- **Screenshots** of failed sends are saved to `screenshots/`
- **Detailed logs** are in `logs/campaign_YYYY-MM-DD.log`
- Rows with send errors are marked `Not Found` in Email Status

If Outlook selectors break (Outlook UI changes):
→ Update selector values in `config.py` under `# Selectors — Outlook Web`

---

## ⚠️ Notes

- **Outlook type**: Uses **Outlook Web** (`outlook.office.com`).  
  Desktop Outlook is a Win32 app — Playwright cannot automate it.
- **Login**: You log in manually once per session. The browser session is reused for all emails.
- **No CC/BCC**: Each email is sent individually with no CC or BCC.
- **Template used**: Only Template 1 (first row of Sheet2) is used. Change `template_index` in `main.py` to use others.
