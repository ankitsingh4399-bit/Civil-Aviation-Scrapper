# Civil Aviation Website Scraper

Scrapes headline metrics (Domestic Traffic, International Traffic, On-Time Performance, Passenger Load Factor) from the **Ministry of Civil Aviation, India** website home page and emails a formatted digest via Outlook.

## ✨ What it does
- Requests `https://www.civilaviation.gov.in` (falls back to `http://` if `https://` fails).
- Parses specific metric blocks using **BeautifulSoup**.
- Cleans text, removes Devanagari-prefixed strings, and pairs labels with values.
- Composes an **Outlook email** to `skdteam@goindigo.in` with the extracted metrics and a timestamp.

## 📁 Suggested Structure
```
project-root/
├─ civil_aviation_scrapper.py
└─ requirements.txt
```

## 🛠️ Requirements
- **Operating System:** Windows (required for Outlook COM via `pywin32`).
- **Python:** 3.9–3.11
- **Dependencies:** `beautifulsoup4`, `requests`, `pywin32`

Install:
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```
Example `requirements.txt`:
```txt
beautifulsoup4>=4.12
requests>=2.31
pywin32>=306
```

## 🚀 Usage
Run from anywhere:
```bash
python civil_aviation_scrapper.py
```
On success, an Outlook email with the current metrics is sent to `skdteam@goindigo.in`. The subject line includes the run timestamp (e.g., `Civil Aviation Website Update: 18-Dec-25 12:02`).

## ⚙️ How it works (positions & selectors)
- Uses the following **CSS class containers** on the homepage:
  - `div.views-element-container.col-lg-4.col-md-6.col-sm-12.domestic-traffic`
  - `div.views-element-container.col-lg-4.col-md-6.col-sm-12.international-traffic`
  - `div.views-element-container.col-lg-4.col-md-6.col-sm-12.on-time-performance`
  - `div.views-element-container.col-lg-4.col-md-6.col-sm-12.passenger-load-factor`
- For each block, splits text by newline, trims, removes empty strings, and filters out entries starting with the **Devanagari Unicode range** `U+0900–U+097F`.
- Pairs items `[label, value]` into `"label: value"` strings and joins with line breaks.

## 🧪 Testing Tips
- Run once with Outlook open and signed-in.
- Print intermediate results (`domestic`, `international`, `otp`, `plf`) to verify parsing.
- If parsing fails, inspect the live HTML structure (classes may change). Consider switching to more robust selectors.

## 🔐 Notes & Best Practices
- **Site changes**: The homepage layout/classes may change—add error handling around `.find()` calls and guard against `None`.
- **Network**: A custom `User-Agent` is set; consider adding retry/backoff for transient failures.
- **Email**: Externalize recipients via env vars (`NOTIFY_TO`) and read via `os.environ`.
- **Logging**: Add structured logging for operations and exceptions.

## ⏱️ Scheduling (Windows Task Scheduler)
- Trigger this script daily or weekly to capture updates.
- Run only when Outlook is available.

## 📄 License
Proprietary/Internal (IndiGo). Update if needed.

## 👤 Author
Divyanshu Upadhyay (PRM, ISC)
