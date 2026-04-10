from bs4 import BeautifulSoup
import requests
import re
import warnings
import win32com.client
import sys
import time
from datetime import datetime
import unicodedata

warnings.filterwarnings("ignore")

# ============================================================
# CONFIG (EDIT THESE)
# ============================================================
HOME_URLS = [
    "https://www.civilaviation.gov.in",
    "http://www.civilaviation.gov.in",
]

NOTIFY_RECIPIENTS = "skdteam@goindigo.in;indigoslots@goindigo.in"
ERROR_RECIPIENTS = "ankit.singh8@goindigo.in"

REQUEST_TIMEOUT = 30
RETRY_LIMIT = 4
RETRY_SLEEP_SECONDS = 4

# Use system/environment proxy settings
TRUST_ENV_PROXY = True

# If your proxy must be explicitly set (ask IT for the proxy URL):
PROXIES = None
# Example:
# PROXIES = {
#     "http":  "http://proxy.company.com:8080",
#     "https": "http://proxy.company.com:8080",
# }

# If your corporate proxy requires Windows Integrated Auth:
# pip install requests-negotiate-sspi
USE_WINDOWS_PROXY_AUTH = True

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Encoding": "gzip, deflate",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "DNT": "1",
    "Connection": "close",
    "Upgrade-Insecure-Requests": "1",
}

DEBUG = True  # set False once stable


# ============================================================
# PARSING HELPERS
# ============================================================
def sublists_between_hindi_word_markers(
    lst,
    include_pre_first=False,
    include_post_last=False,
    skip_empty=False,
    normalize_unicode=True,
):
    def norm(s):
        return unicodedata.normalize("NFC", s) if normalize_unicode else s

    devanagari_start_re = re.compile(r"^\s*[\u0900-\u097F]")

    def starts_with_hindi_word(item) -> bool:
        s = norm(str(item))
        return bool(devanagari_start_re.match(s))

    result = []
    last_marker_idx = None

    for i, item in enumerate(lst):
        if starts_with_hindi_word(item):
            if last_marker_idx is None:
                if include_pre_first and i > 0:
                    seg = lst[:i]
                    if not (skip_empty and len(seg) == 0):
                        result.append(seg)
            else:
                seg = lst[last_marker_idx + 1 : i]
                if not (skip_empty and len(seg) == 0):
                    result.append(seg)
            last_marker_idx = i

    # ✅ Key fix: keep segment after last marker (often the 6th tile)
    if include_post_last:
        seg = lst[:] if last_marker_idx is None else lst[last_marker_idx + 1 :]
        if not (skip_empty and len(seg) == 0):
            result.append(seg)

    return result


def join_elements_pattern(list_of_lists):
    out = []
    for sub in list_of_lists:
        if not sub:
            continue

        sub_strs = [str(x).strip() for x in sub if str(x).strip()]

        if len(sub_strs) > 2:
            base = f"{sub_strs[0]}: {sub_strs[1]}"
            for elem in sub_strs[2:]:
                out.append(f"{base} {elem}")
        elif len(sub_strs) == 2:
            out.append(f"{sub_strs[0]}: {sub_strs[1]}")
    return out


def clean_lines(text: str):
    return [ln.strip() for ln in text.split("\n") if ln and ln.strip()]


# ============================================================
# EMAIL
# ============================================================
def error_email(date_time, error):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = ERROR_RECIPIENTS
    mail.Subject = f"Civil Aviation Update Error: {date_time}"
    mail.HTMLBody = f"""
    <html>
      <body>
        <p>Dear Team,</p>
        <p>Error <b>{error}</b> occurred while running the script.</p>
        <br><br><br>
        <p>Thanks &amp; Regards,<br><br>Publication Team<br>(Comm,ISC)</p>
      </body>
    </html>
    """
    mail.Send()
    print("\nError Email sent!\n")


def notification_email(date_time, domestic, international, otp, plf):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = NOTIFY_RECIPIENTS
    mail.Subject = f"Civil Aviation Website Update: {date_time}"
    mail.HTMLBody = f"""
    <html>
      <body>
        <p>Dear All,</p>
        <p>Please find the below data as of today from Civil Aviation website:</p>
        <br>

        <p><b>{domestic}</b></p>
        <br>

        <p><b>{international}</b></p>
        <br>

        <p><b>{otp}</b></p>
        <br>

        <p><b>{plf}</b></p>

        <br><br><br>
        <p>Thanks &amp; Regards,<br><br>Publication Team<br>(Comm,ISC)</p>
      </body>
    </html>
    """
    mail.Send()
    print("\nNotification Email sent successfully!\n")


# ============================================================
# NETWORK
# ============================================================
def build_session():
    s = requests.Session()
    s.headers.update(HEADERS)
    s.trust_env = TRUST_ENV_PROXY

    if PROXIES:
        s.proxies.update(PROXIES)

    if USE_WINDOWS_PROXY_AUTH:
        try:
            from requests_negotiate_sspi import HttpNegotiateAuth
            s.auth = HttpNegotiateAuth()
        except ImportError:
            raise RuntimeError(
                "Missing dependency: requests-negotiate-sspi\n"
                "Install with: pip install requests-negotiate-sspi\n"
                "Or set USE_WINDOWS_PROXY_AUTH=False"
            )

    return s


def fetch_homepage(session: requests.Session):
    last_err = None
    for url in HOME_URLS:
        try:
            r = session.get(url, timeout=REQUEST_TIMEOUT)
            r.raise_for_status()
            if DEBUG:
                print("[DEBUG] fetched:", r.url, "status:", r.status_code)
                print("[DEBUG] content-type:", r.headers.get("content-type"))
            return r.text
        except Exception as e:
            last_err = e
    raise RuntimeError(f"Unable to fetch homepage. Last error: {last_err}")


# ============================================================
# SCRAPING
# ============================================================
def extract_section(soup: BeautifulSoup, unique_class: str, raw_html: str):
    # Robust selector: only the unique class token
    div = soup.select_one(f"div.{unique_class}")
    if not div:
        if DEBUG:
            title = soup.title.string.strip() if soup.title and soup.title.string else "No title"
            print(f"[DEBUG] '{unique_class}' NOT found. Page title:", title)
            print(f"[DEBUG] '{unique_class}' string present in HTML?", unique_class in raw_html)
        return f"{unique_class}: Not found"

    raw_text = div.get_text("\n", strip=True)
    lines = clean_lines(raw_text)

    grouped = sublists_between_hindi_word_markers(
        lines,
        include_post_last=True,   # ✅ keep last tile
        skip_empty=True
    )
    formatted = join_elements_pattern(grouped)
    if not formatted:
        formatted = lines

    return "<br>".join(formatted)


def run_scrapper():
    try:
        session = build_session()
        html = fetch_homepage(session)
        soup = BeautifulSoup(html, "html.parser")

        domestic = extract_section(soup, "domestic-traffic", html)
        international = extract_section(soup, "international-traffic", html)
        otp = extract_section(soup, "on-time-performance", html)
        plf = extract_section(soup, "passenger-load-factor", html)

        now = datetime.now().strftime("%d-%b-%y %H:%M")
        notification_email(now, domestic, international, otp, plf)
        return 1

    except Exception as e:
        now = datetime.now().strftime("%d-%b-%y %H:%M")
        print(f"Error {e} Occurred!")
        try:
            error_email(now, e)
        except Exception as mail_err:
            print("Additionally failed to send error email:", mail_err)
        return 0


if __name__ == "__main__":
    x = 0
    i = 0
    while x != 1:
        if i < RETRY_LIMIT:
            i += 1
            x = run_scrapper()
            if x != 1:
                time.sleep(RETRY_SLEEP_SECONDS)
        else:
            now = datetime.now().strftime("%d-%b-%y %H:%M")
            error_email(now, f"Limit ({RETRY_LIMIT}) reached for number of iterations")
            sys.exit()