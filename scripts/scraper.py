"""
WebPT Scheduler History scraper — GitHub Actions edition.

Same logic/fixes as the desktop version (robust Custom Range click,
no-data detection, session-bounce recovery) but:
  - reads START_DATE / END_DATE / WEBPT_USERNAME / WEBPT_PASSWORD from
    environment variables (set by the GitHub Actions workflow) instead
    of hardcoded constants
  - runs Chrome headless
  - writes output to ./output/raw relative to the repo checkout instead
    of a C:\ path, since the runner filesystem is ephemeral
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
import time
import os
import sys
import traceback
from datetime import datetime
import pandas as pd
import glob

# ══════════════════════════════════════════════════════════
#  CONFIG — from environment (set by the workflow / job)
# ══════════════════════════════════════════════════════════
START_DATE = os.environ.get("START_DATE")   # MM/DD/YYYY
END_DATE   = os.environ.get("END_DATE")     # MM/DD/YYYY
USERNAME   = os.environ.get("WEBPT_USERNAME")
PASSWORD   = os.environ.get("WEBPT_PASSWORD")
MAX_RETRIES = int(os.environ.get("MAX_RETRIES") or "6")
# Debug aid: set CLINIC_LIMIT=1 (or any small N) via workflow input/env to
# only process the first N clinics — for cheaply testing a fix like this
# one instead of burning a full 35-clinic run per attempt.
CLINIC_LIMIT = int(os.environ.get("CLINIC_LIMIT") or "0") or None
# User-facing feature (not debug): restrict the scrape to specific named
# clinics, e.g. from the web app's clinic multi-select. Pipe-delimited
# since clinic names can't contain "|" but a couple do contain commas
# in principle, so comma-joining would be unsafe.
CLINIC_NAMES_RAW = os.environ.get("CLINIC_NAMES", "").strip()
CLINIC_NAMES = [c.strip() for c in CLINIC_NAMES_RAW.split("|") if c.strip()] if CLINIC_NAMES_RAW else None
WEBPT_URL = "https://app.webpt.com"

if not START_DATE or not END_DATE:
    print("❌ START_DATE and END_DATE must be set (env vars, MM/DD/YYYY).")
    sys.exit(1)
if not USERNAME or not PASSWORD:
    print("❌ WEBPT_USERNAME and WEBPT_PASSWORD must be set (from GitHub secrets).")
    sys.exit(1)

REPO_ROOT    = os.environ.get("GITHUB_WORKSPACE", os.getcwd())
OUTPUT_DIR   = os.path.join(REPO_ROOT, "output", "raw")
DOWNLOAD_DIR = os.path.join(REPO_ROOT, "output", "downloads_temp")
for d in [OUTPUT_DIR, DOWNLOAD_DIR]:
    os.makedirs(d, exist_ok=True)

driver = None


class SessionBounced(Exception):
    pass


# ── LOGGING / TIMING ──
def log(msg):
    print(msg, flush=True)

def timer_start():
    return time.time()

def timer_end(start, label=""):
    elapsed = time.time() - start
    log(f"   ⏱️  {label} took {elapsed:.1f}s")
    return elapsed

def start_task(name):
    print(f"\n{'='*55}\n  STARTING: {name}\n{'='*55}")

def finish_task(success=True, output_file=None, error=None):
    print(f"\n{'='*55}")
    if success:
        print(f"   ✅ FINISHED SUCCESSFULLY")
        if output_file:
            print(f"   📁 Output: {output_file}")
    else:
        print(f"   ❌ FAILED: {error}")
    print(f"{'='*55}\n")


# ── BROWSER (headless for CI) ──
def open_browser():
    global driver
    t = timer_start()
    log("🌐 Opening headless Chrome...")
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    # Headless Chrome needs this CDP command for downloads to actually land on disk
    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": DOWNLOAD_DIR,
        })
    except Exception as e:
        log(f"   ⚠️ Could not set CDP download behavior: {e}")
    timer_end(t, "Browser open")
    log("✅ Chrome opened (headless)")

def wait_click(by, value, timeout=30, desc=""):
    log(f"   ⏳ Clicking: {desc or value}")
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.4)
    driver.execute_script("arguments[0].click();", el)
    log(f"   ✅ Clicked: {desc or value}")
    time.sleep(0.8)
    return el

def wait_visible(by, value, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((by, value)))

def clear_type(el, text):
    driver.execute_script("arguments[0].click();", el)
    driver.execute_script("arguments[0].focus();", el)
    time.sleep(0.3)
    el.send_keys(Keys.CONTROL + "a")
    el.send_keys(Keys.DELETE)
    time.sleep(0.2)
    el.send_keys(text)
    time.sleep(0.3)


# ── SESSION / LOGIN-BOUNCE DETECTION ──
def is_on_login_page():
    try:
        return bool(driver.find_elements(By.ID, "username") or driver.find_elements(By.ID, "password"))
    except Exception:
        return False

def recover_session_if_needed(step_name=""):
    if is_on_login_page():
        log(f"   ⚠️ Bounced to LOGIN page after '{step_name}'. Re-authenticating...")
        login()
        go_to_scheduler_history()
        raise SessionBounced(f"Session lost after {step_name}; re-authenticated.")
    try:
        cur_url = driver.current_url
    except Exception:
        cur_url = ""
    if "/scheduler/history" not in cur_url:
        log(f"   ⚠️ Unexpected navigation after '{step_name}' (now at: {cur_url}). Recovering...")
        if is_on_login_page():
            login()
        go_to_scheduler_history()
        raise SessionBounced(f"Navigated away after {step_name}; recovered.")


# ── LOGIN ──
def login():
    t = timer_start()
    log("\n=== LOGIN ===")
    driver.get(WEBPT_URL)
    time.sleep(5)
    try: f = wait_visible(By.ID, "username")
    except: f = wait_visible(By.XPATH, "//input[@type='text' or @type='email']")
    clear_type(f, USERNAME)
    try: wait_click(By.CSS_SELECTOR, "button[type='submit']", desc="Continue")
    except: wait_click(By.XPATH, "//button[contains(text(),'Continue')]", desc="Continue")
    time.sleep(4)
    try: f = wait_visible(By.ID, "password")
    except: f = wait_visible(By.XPATH, "//input[@type='password']")
    clear_type(f, PASSWORD)
    try: wait_click(By.CSS_SELECTOR, "button[type='submit']", desc="Sign In")
    except: wait_click(By.XPATH, "//button[contains(text(),'Sign')]", desc="Sign In")
    time.sleep(6)
    time.sleep(3)
    try:
        oust = WebDriverWait(driver, 7).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.eviction-option.ok")))
        driver.execute_script("arguments[0].click();", oust)
        time.sleep(5)
    except:
        log("   ℹ️ No eviction page")
    timer_end(t, "Login")
    log("✅ Logged in")


# ── NAVIGATION ──
def go_to_scheduler_history(extra_wait=0, max_relogin=2):
    t = timer_start()
    for attempt in range(max_relogin + 1):
        log("   🚀 Navigating to Scheduler History URL...")
        driver.get(f"{WEBPT_URL}/scheduler/history")
        time.sleep(2)
        if is_on_login_page():
            log("   ⚠️ Bounced to login while navigating — re-authenticating...")
            login()
            continue
        try:
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "ClinicChange")))
            WebDriverWait(driver, 30).until(
                lambda d: len(d.find_elements(By.CSS_SELECTOR, "#ClinicChange option")) > 0
            )
            base_wait = 4 + extra_wait
            log(f"   ⏳ Page buffer wait: {base_wait}s...")
            time.sleep(base_wait)
            timer_end(t, "Navigation")
            return True
        except TimeoutException:
            if is_on_login_page():
                log("   ⚠️ Bounced to login while waiting for clinic list — re-authenticating...")
                login()
                continue
            raise
    raise SessionBounced("Could not reach Scheduler History after repeated re-login attempts.")


# ── CLINIC LIST ──
def get_all_clinics():
    log("\n=== FETCHING AVAILABLE CLINICS ===")
    t = timer_start()
    WebDriverWait(driver, 30).until(
        lambda d: len(d.find_elements(By.CSS_SELECTOR, "#ClinicChange option")) > 0
    )
    options_data = driver.execute_script("""
        var sel = document.getElementById('ClinicChange');
        if (!sel) return [];
        var results = [];
        for (var i = 0; i < sel.options.length; i++) {
            results.push(sel.options[i].text.trim());
        }
        return results;
    """)
    excluded = ["Physical Therapy of New York", "Physical Therapy of The City P.C."]
    clinic_names = [c for c in options_data if c and c not in excluded]
    timer_end(t, "Clinic fetch")
    log(f"📋 Found {len(clinic_names)} distinct clinics to query.")
    return clinic_names


# ── CLINIC SWITCH ──
def switch_clinic(clinic_name, retry_num=1):
    log(f"\n🏢 Switching to: '{clinic_name}' (attempt {retry_num})")
    t = timer_start()
    if retry_num > 1:
        extra = (retry_num - 1) * 3
        time.sleep(extra)
    option_value = driver.execute_script("""
        var sel = document.getElementById('ClinicChange');
        if (!sel) return null;
        for (var i = 0; i < sel.options.length; i++) {
            if (sel.options[i].text.trim() === arguments[0]) return sel.options[i].value;
        }
        return null;
    """, clinic_name)
    if not option_value:
        log(f"   ⚠️ Could not find value for clinic '{clinic_name}' — skipping.")
        return
    btn = driver.find_element(By.CSS_SELECTOR, "#ClinicChange_chosen a.chosen-single")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    time.sleep(0.5)
    ActionChains(driver).move_to_element(btn).click().perform()
    time.sleep(2)
    chosen_div = driver.find_element(By.ID, "ClinicChange_chosen")
    classes = chosen_div.get_attribute("class")
    if "chosen-container-active" in classes:
        all_li = driver.find_elements(By.CSS_SELECTOR, "#ClinicChange_chosen .chosen-results li.active-result")
        clicked = False
        for li in all_li:
            if li.text.strip() == clinic_name:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", li)
                time.sleep(0.3)
                ActionChains(driver).move_to_element(li).click().perform()
                clicked = True
                break
        if not clicked:
            _js_trigger_clinic_change(option_value)
    else:
        _js_trigger_clinic_change(option_value)
    reload_wait = 8 + ((retry_num - 1) * 4)
    time.sleep(reload_wait)
    timer_end(t, f"Switch clinic '{clinic_name}'")

def _js_trigger_clinic_change(option_value):
    driver.execute_script("""
        var sel = document.getElementById('ClinicChange');
        if (!sel) return 'NO_SELECT';
        for (var i = 0; i < sel.options.length; i++) {
            if (sel.options[i].value === arguments[0]) { sel.selectedIndex = i; break; }
        }
        sel.removeAttribute('disabled');
        try {
            changeClinic.change(564697, arguments[0]);
            return 'OK_DIRECT_CALL';
        } catch(e) {
            var evt = document.createEvent('HTMLEvents');
            evt.initEvent('change', true, true);
            sel.dispatchEvent(evt);
            return 'OK_EVENT: ' + e.toString();
        }
    """, option_value)


# ── POPOVER HELPERS ──
def _close_any_open_popover():
    try: driver.execute_script("document.body.click();")
    except: pass
    try: ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    except: pass
    time.sleep(0.5)

def _get_visible_popover(timeout=10):
    end = time.time() + timeout
    while time.time() < end:
        popovers = driver.find_elements(By.CSS_SELECTOR, ".popover-content")
        visible = [p for p in popovers if p.is_displayed()]
        if len(visible) == 1:
            return visible[0]
        if len(visible) > 1:
            return visible[-1]
        time.sleep(0.3)
    return None

def _open_date_popover(max_attempts=4):
    selectors = [
        (By.CSS_SELECTOR, "input.bn-textfield-input-textbox",           "Date Field input box"),
        (By.CSS_SELECTOR, "button[data-test-input-icon-button='true']", "Calendar icon button"),
    ]
    _close_any_open_popover()
    for attempt in range(max_attempts):
        by, selector, desc = selectors[attempt % len(selectors)]
        try:
            wait_click(by, selector, timeout=10, desc=desc)
        except Exception as e:
            log(f"   ⚠️ Could not click {desc} (attempt {attempt+1}): {e}")
            continue
        popover = _get_visible_popover(timeout=6)
        if popover is not None:
            return popover
        log(f"   ⚠️ Popover didn't appear after {desc} (attempt {attempt+1}) — retrying with the other trigger...")
        _close_any_open_popover()
    return None


# ── CUSTOM RANGE TAB (text-based, verified) ──
def _click_custom_range_tab(popover, max_attempts=4):
    for attempt in range(max_attempts):
        try:
            candidates = popover.find_elements(By.XPATH, ".//*[contains(normalize-space(text()),'Custom Range')]")
            target = next((c for c in candidates if c.is_displayed()), None)
            if target is None:
                log(f"   ⚠️ 'Custom Range' text not found (attempt {attempt+1}) — reopening picker...")
                popover = _open_date_popover()
                if popover is None:
                    return None
                continue
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target)
            time.sleep(0.3)
            driver.execute_script("arguments[0].click();", target)
            time.sleep(1)
            try:
                WebDriverWait(driver, 5).until(
                    lambda d: len(popover.find_elements(By.CSS_SELECTOR, ".DayPicker-Day")) > 0
                )
                log("   ✅ 'Custom Range' confirmed — calendar day grid visible.")
                return popover
            except TimeoutException:
                log(f"   ⚠️ Clicked but calendar grid never appeared (attempt {attempt+1}) — retrying.")
                continue
        except StaleElementReferenceException:
            popover = _get_visible_popover(timeout=5) or _open_date_popover()
            if popover is None:
                return None
            continue
    return None


def _click_calendar_day(popover, aria_label, mode):
    target_dt = datetime.strptime(aria_label, "%a %b %d %Y")
    for _ in range(24):
        try:
            day_el = popover.find_element(
                By.CSS_SELECTOR,
                f".DayPicker-Day[aria-label='{aria_label}']:not(.DayPicker-Day--outside)"
            )
            if day_el.is_displayed():
                driver.execute_script("arguments[0].click();", day_el)
                log(f"   ✅ Chosen {mode} target square: {aria_label}")
                return True
        except:
            pass
        captions = popover.find_elements(By.CSS_SELECTOR, ".DayPicker-Caption div")
        months = []
        for cap in captions:
            try: months.append(datetime.strptime(cap.text.strip(), "%B %Y"))
            except: pass
        if not months:
            break
        max_m, min_m = max(months), min(months)
        target_m = datetime(target_dt.year, target_dt.month, 1)
        if target_m > max_m:
            nav = popover.find_element(By.CSS_SELECTOR, ".DayPicker-NavButton--next")
            driver.execute_script("arguments[0].click();", nav)
        elif target_m < min_m:
            nav = popover.find_element(By.CSS_SELECTOR, ".DayPicker-NavButton--prev")
            driver.execute_script("arguments[0].click();", nav)
        else:
            time.sleep(0.5)
        time.sleep(0.8)
    log(f"   ❌ Could not find/click {mode} day '{aria_label}'.")
    return False


def set_date_range(start_date, end_date, retry_num=1):
    t = timer_start()
    log(f"   📅 Setting Date Range: {start_date} → {end_date}")
    start_dt = datetime.strptime(start_date, "%m/%d/%Y")
    end_dt   = datetime.strptime(end_date,   "%m/%d/%Y")
    start_label = start_dt.strftime("%a %b %d %Y")
    end_label   = end_dt.strftime("%a %b %d %Y")

    popover = _open_date_popover()
    if popover is None:
        timer_end(t, "Set date range (failed to open)")
        return False
    popover = _click_custom_range_tab(popover)
    if popover is None:
        timer_end(t, "Set date range (custom range failed)")
        return False
    if not _click_calendar_day(popover, start_label, "start"):
        timer_end(t, "Set date range (start day failed)")
        return False
    time.sleep(1.0)
    popover = _get_visible_popover(timeout=5) or popover
    if not _click_calendar_day(popover, end_label, "end"):
        timer_end(t, "Set date range (end day failed)")
        return False
    time.sleep(2)
    timer_end(t, "Set date range")
    return True


def _dump_diagnostics(label):
    """
    Saves the current page HTML + a screenshot to output/debug/ so we can
    see EXACTLY what's on the page when something goes wrong, instead of
    guessing at another selector blind. Never raises — best effort only.
    """
    try:
        debug_dir = os.path.join(REPO_ROOT, "output", "debug")
        os.makedirs(debug_dir, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        html_path = os.path.join(debug_dir, f"{label}_{ts}.html")
        png_path = os.path.join(debug_dir, f"{label}_{ts}.png")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        driver.save_screenshot(png_path)
        log(f"   🐞 Diagnostics saved: {os.path.basename(html_path)}, {os.path.basename(png_path)}")
    except Exception as e:
        log(f"   ⚠️ Could not save diagnostics: {e}")


# ── GENERATE ──
def click_generate(retry_num=1):
    t = timer_start()
    log("   🖱️ Clicking Generate...")
    app_wait = 60 + ((retry_num - 1) * 30)
    try:
        WebDriverWait(driver, app_wait).until(
            EC.presence_of_element_located((By.ID, "app-scheduler-change-history-root"))
        )
    except TimeoutException:
        if is_on_login_page():
            timer_end(t, "Generate (session lost)")
            return "SESSION_LOST"
        timer_end(t, "Generate (app never loaded)")
        return "FAILED"

    time.sleep(1 + (retry_num - 1))
    # Case-insensitive text match via translate() (XPath 1.0 has no lower-case()).
    # ALSO tries ANY button/clickable element mentioning "generate", not just
    # inside the app root, in case the container id itself isn't what we think.
    LOWER = "abcdefghijklmnopqrstuvwxyz"
    UPPER = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    ci = f"translate(normalize-space(.), '{UPPER}', '{LOWER}')"
    btn_wait = 15 + ((retry_num - 1) * 10)
    generated = False
    text_xpaths = [
        f"//div[@id='app-scheduler-change-history-root']//button[contains({ci},'generate')]",
        f"//button[contains({ci},'generate')]",
        f"//*[self::button or self::a or self::div[@role='button']][contains({ci},'generate')]",
    ]
    for xpath in text_xpaths:
        try:
            btn = WebDriverWait(driver, btn_wait).until(EC.element_to_be_clickable((By.XPATH, xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", btn)
            log(f"   ✅ Generate clicked via: {xpath}")
            generated = True
            break
        except TimeoutException:
            continue
    if not generated:
        _dump_diagnostics("generate_button_not_found")
        timer_end(t, "Generate (failed)")
        return "FAILED"

    data_wait = 360 + ((retry_num - 1) * 120)
    poll_interval = 3
    end_time = time.time() + data_wait
    while time.time() < end_time:
        if is_on_login_page():
            timer_end(t, "Generate (session lost)")
            return "SESSION_LOST"
        if driver.find_elements(By.XPATH, "//a[@download='SchedulerHistory.csv']"):
            timer_end(t, "Generate + data load")
            return "DATA_READY"
        if driver.find_elements(By.CSS_SELECTOR, ".AlertMessage") or \
           driver.find_elements(By.XPATH, "//*[contains(normalize-space(text()),'No results found')]"):
            timer_end(t, "Generate (no data)")
            return "NO_DATA"
        time.sleep(poll_interval)
    timer_end(t, "Generate (timeout)")
    return "TIMEOUT"


# ── DOWNLOAD ──
def click_download():
    t = timer_start()
    for f in glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")):
        try: os.remove(f)
        except: pass
    try:
        dll = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@download='SchedulerHistory.csv']"))
        )
        driver.execute_script("arguments[0].click();", dll)
    except Exception as e:
        log(f"   ❌ Failed to click download link: {e}")
        return None
    end = time.time() + 60
    while time.time() < end:
        files = [f for f in glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")) if not f.endswith(".crdownload")]
        if files:
            latest = max(files, key=os.path.getmtime)
            time.sleep(1)
            timer_end(t, "Download")
            return latest
        time.sleep(1)
    log("   ❌ Download timed out after 60s.")
    return None


# ── SINGLE CLINIC ATTEMPT ──
def process_clinic(clinic, attempt):
    log(f"\n   🔄 Attempt {attempt}/{MAX_RETRIES} for '{clinic}'")
    attempt_start = timer_start()

    switch_clinic(clinic, retry_num=attempt)
    recover_session_if_needed("switch_clinic")

    go_to_scheduler_history(extra_wait=(attempt - 1) * 2)
    recover_session_if_needed("go_to_scheduler_history")

    date_ok = set_date_range(START_DATE, END_DATE, retry_num=attempt)
    recover_session_if_needed("set_date_range")
    if not date_ok:
        timer_end(attempt_start, f"Attempt {attempt} total (date range failed)")
        return None

    status = click_generate(retry_num=attempt)
    if status == "SESSION_LOST":
        recover_session_if_needed("click_generate")
    if status == "NO_DATA":
        timer_end(attempt_start, f"Attempt {attempt} total (no data)")
        return pd.DataFrame()
    if status != "DATA_READY":
        timer_end(attempt_start, f"Attempt {attempt} total (not ready)")
        return None

    csv_path = click_download()
    if csv_path:
        df = pd.read_csv(csv_path, on_bad_lines='skip', engine='python')
        os.remove(csv_path)
        timer_end(attempt_start, f"Attempt {attempt} total")
        return df
    timer_end(attempt_start, f"Attempt {attempt} total (no data)")
    return None


# ── MAIN ──
def run_scheduler_history_by_clinic():
    global driver
    start_task("Clinic Scheduler History Scraper Loop (CI)")
    script_start = timer_start()
    try:
        open_browser()
        login()
        go_to_scheduler_history()
        full_clinic_list = get_all_clinics()
        # Publish the FULL discovered clinic list (before any filtering) so the
        # web app's clinic dropdown always stays in sync with what WebPT
        # actually has, regardless of what this particular run was scoped to.
        gh_out_early = os.environ.get("GITHUB_OUTPUT")
        if gh_out_early:
            with open(gh_out_early, "a") as f:
                f.write(f"all_clinics={'|'.join(full_clinic_list)}\n")

        all_clinics = full_clinic_list
        if CLINIC_NAMES:
            wanted = {c.lower() for c in CLINIC_NAMES}
            all_clinics = [c for c in all_clinics if c.lower() in wanted]
            log(f"   🎯 Restricted to {len(all_clinics)} selected clinic(s): {', '.join(all_clinics)}")
        if CLINIC_LIMIT:
            log(f"   🐞 DEBUG MODE: limiting to first {CLINIC_LIMIT} clinic(s) (CLINIC_LIMIT set).")
            all_clinics = all_clinics[:CLINIC_LIMIT]
        batch_results, failed_clinics, no_data_clinics = [], [], []
        if not all_clinics:
            finish_task(success=False, error="No clinics found to query.")
            sys.exit(1)
        for idx, clinic in enumerate(all_clinics, start=1):
            clinic_start = timer_start()
            log(f"\n{'─'*55}\n🏁 Processing ({idx}/{len(all_clinics)}): '{clinic}'\n{'─'*55}")
            df = None
            for attempt in range(1, MAX_RETRIES + 1):
                try:
                    df = process_clinic(clinic, attempt)
                    if df is not None:
                        break
                    if attempt < MAX_RETRIES:
                        time.sleep(attempt * 5)
                except SessionBounced as e:
                    log(f"   ↩️ Session recovered after WebPT bounce ({e}) — retrying clinic.")
                    if attempt < MAX_RETRIES:
                        time.sleep(attempt * 3)
                except Exception as e:
                    log(f"   ❌ Attempt {attempt} exception: {e}")
                    log(traceback.format_exc())
                    if attempt < MAX_RETRIES:
                        time.sleep(attempt * 5)
            elapsed = time.time() - clinic_start
            if df is not None:
                batch_results.append((clinic, df))
                if df.empty:
                    no_data_clinics.append(clinic)
                log(f"   ✅ '{clinic}' done in {elapsed:.1f}s")
            else:
                log(f"   ❌ '{clinic}' FAILED after {MAX_RETRIES} attempts ({elapsed:.1f}s)")
                failed_clinics.append(clinic)
                batch_results.append((clinic, pd.DataFrame()))

        log(f"\n{'='*55}\n📊 SCRAPE SUMMARY\n{'='*55}")
        log(f"   Total clinics:      {len(all_clinics)}")
        log(f"   ✅ With data:        {len(all_clinics) - len(failed_clinics) - len(no_data_clinics)}")
        log(f"   ℹ️ No data (valid):  {len(no_data_clinics)}")
        log(f"   ❌ Failed:           {len(failed_clinics)}")
        if failed_clinics:
            log(f"   Failed list: {', '.join(failed_clinics)}")
        timer_end(script_start, "Total scrape runtime")

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"SchedulerHistory_ByClinic_{ts}.xlsx"
        filepath = os.path.join(OUTPUT_DIR, filename)
        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            for clinic_name, df in batch_results:
                clean_tab = clinic_name[:30].replace("[", "").replace("]", "")
                if not df.empty:
                    df.to_excel(writer, index=False, sheet_name=clean_tab)
                else:
                    pd.DataFrame({"Note": ["No data returned."]}).to_excel(writer, index=False, sheet_name=clean_tab)
        finish_task(success=True, output_file=filepath)

        # Hand the raw output path to the next pipeline step via GITHUB_OUTPUT
        gh_out = os.environ.get("GITHUB_OUTPUT")
        if gh_out:
            with open(gh_out, "a") as f:
                f.write(f"raw_file={filepath}\n")

        # Fail the job if EVERY clinic failed outright (as opposed to legitimately empty)
        if failed_clinics and len(failed_clinics) == len(all_clinics):
            sys.exit(1)
        return filepath
    except Exception as e:
        log(f"❌ FATAL ERROR: {e}")
        log(traceback.format_exc())
        finish_task(success=False, error=str(e))
        sys.exit(1)
    finally:
        if driver:
            driver.quit()
            log("🔒 Browser closed.")


if __name__ == "__main__":
    run_scheduler_history_by_clinic()
