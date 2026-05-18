"""
JCtrans Company Directory Scraper
=================================

Scrape company directory from https://www.jctrans.com/en/directory/.

Usage::

    python scraper.py                         # Run for all countries
    python scraper.py --country Japan         # Run for a single country
    python scraper.py --country Japan India   # Run for several countries
    python scraper.py --resume                # Skip countries already done

Output Excel layout (one row per contact):
    A  STT (running index)
    B  Country
    C  Company name
    D  Member ID
    E  Year
    F  Location
    G  Website
    H  Main Business        (chips joined by " | ")
    I  Sea Freight Adv.     (chips joined by " | ")
    J  Air Freight Adv.     (chips joined by " | ")
    K  Contact name
    L  Position
    M  Email
    N  Phone
    O  WeChat
    P  WhatsApp
    Q  Skype

Suspended / non-member companies (red badge "It is NOT a JCtrans member" or
"MEMBERSHIP SUSPEND") are skipped entirely per requirement.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import random
import re
import sys
import time
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# ──────────────────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────────────────

LOGIN_EMAIL = os.environ.get("JCT_EMAIL", "brent@pio-logistics.vn")
LOGIN_PASSWORD = os.environ.get("JCT_PASSWORD", "0925587314aA")

DIRECTORY_URL = "https://www.jctrans.com/en/directory/"

EXCEL_PATH = Path(os.environ.get("JCT_EXCEL", "jctrans_data.xlsx"))
SHEET_NAME = "Data"
CHECKPOINT_PATH = Path(os.environ.get("JCT_CHECKPOINT", "jctrans_checkpoint.json"))
LOG_DIR = Path("logs")
LOG_DIR.mkdir(exist_ok=True)

HEADERS = [
    "STT",                  # A
    "Country",              # B
    "Company name",         # C
    "Member ID",            # D
    "Year",                 # E
    "Location",             # F
    "Website",              # G
    "Main Business",        # H
    "Sea Freight Adv.",     # I
    "Air Freight Adv.",     # J
    "Contact name",         # K
    "Position",              # L
    "Email",                # M
    "Phone",                # N
    "WeChat",               # O
    "WhatsApp",             # P
    "Skype",                # Q
]

# Icon id (inside <g id="icon/XYZ">) → column index in HEADERS (0-based)
ICON_TO_COL = {
    "mail": 12,       # M  Email
    "phone": 13,      # N  Phone
    "wechat": 14,     # O  WeChat
    "whatsapp": 15,   # P  WhatsApp
    "skype": 16,      # Q  Skype
}

DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36"
)


# ──────────────────────────────────────────────────────────────────────────
# Logging
# ──────────────────────────────────────────────────────────────────────────

logger = logging.getLogger("jct")


def setup_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    fmt = "%(asctime)s %(levelname)-7s %(message)s"
    datefmt = "%H:%M:%S"
    logging.basicConfig(level=level, format=fmt, datefmt=datefmt)
    fh = logging.FileHandler(LOG_DIR / f"scrape_{datetime.now():%Y%m%d_%H%M%S}.log", encoding="utf-8")
    fh.setFormatter(logging.Formatter(fmt, datefmt))
    logging.getLogger().addHandler(fh)


# ──────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────


def make_driver(headless: bool = True) -> WebDriver:
    """Create a Chrome WebDriver. Headless by default (only supported mode on CI)."""
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--lang=en-US")
    opts.add_argument(f"--user-agent={DEFAULT_USER_AGENT}")
    chrome_bin = os.environ.get("CHROME_BIN")
    if chrome_bin and Path(chrome_bin).exists():
        opts.binary_location = chrome_bin
    elif Path("/home/ubuntu/.local/bin/google-chrome").exists():
        opts.binary_location = "/home/ubuntu/.local/bin/google-chrome"
    driver = webdriver.Chrome(options=opts)
    driver.set_page_load_timeout(180)
    return driver


def safe_text(el: WebElement) -> str:
    try:
        return (el.text or "").strip()
    except StaleElementReferenceException:
        return ""


def safe_get(driver: WebDriver, url: str, settle: float = 6.0) -> None:
    """Navigate to URL and wait a bit for SPA to settle."""
    last_exc: Exception | None = None
    for attempt in range(3):
        try:
            driver.get(url)
            time.sleep(settle)
            return
        except (TimeoutException, WebDriverException) as exc:
            last_exc = exc
            logger.warning("page load failed (attempt %s): %s", attempt + 1, exc)
            time.sleep(5)
    if last_exc:
        raise last_exc


def jitter(a: float = 0.4, b: float = 1.2) -> None:
    time.sleep(random.uniform(a, b))


def close_blocking_overlays(driver: WebDriver) -> None:
    """Try to close ad / popup overlays that block clicks."""
    selectors = [
        "div.el-dialog__headerbtn",        # generic Element-Plus close
        ".el-dialog .el-icon-close",
        ".close, .icon-close",
    ]
    for sel in selectors:
        for el in driver.find_elements(By.CSS_SELECTOR, sel):
            try:
                if el.is_displayed():
                    el.click()
                    time.sleep(0.4)
            except Exception:
                pass


# ──────────────────────────────────────────────────────────────────────────
# Login
# ──────────────────────────────────────────────────────────────────────────


def click_confirm_logged_elsewhere(driver: WebDriver) -> bool:
    """If a 'Tips: account is logged in elsewhere' dialog is shown, click Confirm."""
    candidates = driver.find_elements(
        By.XPATH,
        "//div[contains(@class,'el-message-box') or contains(@class,'el-dialog')]"
        "//button[.//span[contains(.,'Confirm') or contains(.,'OK')]]",
    )
    for btn in candidates:
        try:
            if btn.is_displayed():
                logger.info("clicking Confirm on 'account logged elsewhere'")
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(3)
                return True
        except Exception:
            pass
    return False


def login(driver: WebDriver, email: str, password: str) -> None:
    """Open directory page, open Sign-In dialog, fill credentials, submit."""
    safe_get(driver, DIRECTORY_URL, settle=6)
    wait = WebDriverWait(driver, 30)

    # If already logged in (no SIGN IN button), skip
    signin_buttons = [b for b in driver.find_elements(By.CSS_SELECTOR, "button.login-btn") if b.is_displayed()]
    if not signin_buttons:
        logger.info("appears already logged in")
        return

    # Open dialog
    btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.login-btn")))
    driver.execute_script("arguments[0].click();", btn)
    time.sleep(2)

    # Fill credentials
    user_input = wait.until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, ".form-user input.el-input__inner"))
    )
    pass_input = wait.until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, ".password-item input.el-input__inner"))
    )
    user_input.clear()
    user_input.send_keys(email)
    pass_input.clear()
    pass_input.send_keys(password)
    time.sleep(0.4)

    # Submit
    submit_btn = driver.find_element(By.CSS_SELECTOR, "button.submit-btn")
    driver.execute_script("arguments[0].click();", submit_btn)
    time.sleep(6)

    # Handle "account is logged in" dialog if present
    click_confirm_logged_elsewhere(driver)
    time.sleep(4)
    close_blocking_overlays(driver)

    # Verify
    leftover = [b for b in driver.find_elements(By.CSS_SELECTOR, "button.login-btn") if b.is_displayed()]
    if leftover:
        raise RuntimeError("Login appears to have failed: SIGN IN button still visible")
    logger.info("login OK")


# ──────────────────────────────────────────────────────────────────────────
# Country dropdown
# ──────────────────────────────────────────────────────────────────────────


def get_country_input(driver: WebDriver) -> WebElement:
    """Return the <input> element for the 'All Countries/Regions' filter on directory page."""
    inputs = driver.find_elements(By.CSS_SELECTOR, "input.el-select__input")
    for inp in inputs:
        try:
            ph = inp.find_element(
                By.XPATH, "./ancestor::div[contains(@class,'el-select__wrapper')]"
            ).text
        except NoSuchElementException:
            continue
        # Check whether the wrapper contains 'All Countries' placeholder
        try:
            wrapper = inp.find_element(
                By.XPATH, "./ancestor::div[contains(@class,'el-select')][1]"
            )
            if "All Countries" in wrapper.text or "Countries/Regions" in wrapper.text:
                return inp
        except NoSuchElementException:
            continue
        if "All Countries" in (ph or ""):
            return inp
    # Fallback: find by aria-controls of empty country dropdown
    fallback = driver.find_elements(By.XPATH, "//input[@aria-controls and @class='el-select__input']")
    for inp in fallback:
        ac = inp.get_attribute("aria-controls") or ""
        try:
            target_ul = driver.find_element(By.CSS_SELECTOR, f"ul#{ac}")
            panel = target_ul.find_element(
                By.XPATH, "./ancestor::div[contains(@class,'el-select-dropdown')][1]"
            )
            if "company-search-country" in (panel.get_attribute("class") or ""):
                return inp
        except NoSuchElementException:
            continue
    raise RuntimeError("Country input element not found")


def open_country_dropdown(driver: WebDriver) -> list[WebElement]:
    """Click the country filter input and return visible <li> options."""
    inp = get_country_input(driver)
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
    inp.click()
    time.sleep(2)
    items = driver.find_elements(
        By.CSS_SELECTOR, "div.el-select-dropdown.company-search-country li.el-select-dropdown__item"
    )
    if not items:
        # try a longer wait
        time.sleep(3)
        items = driver.find_elements(
            By.CSS_SELECTOR, "div.el-select-dropdown.company-search-country li.el-select-dropdown__item"
        )
    return items


def list_all_countries(driver: WebDriver) -> list[str]:
    items = open_country_dropdown(driver)
    names = []
    for it in items:
        try:
            if it.is_displayed():
                t = safe_text(it)
                if t:
                    names.append(t)
        except Exception:
            pass
    # close dropdown by pressing escape
    try:
        get_country_input(driver).send_keys(Keys.ESCAPE)
    except Exception:
        pass
    time.sleep(0.5)
    return names


def select_country(driver: WebDriver, country_name: str) -> bool:
    items = open_country_dropdown(driver)
    target = None
    for it in items:
        if safe_text(it).lower() == country_name.lower():
            target = it
            break
    if target is None:
        # Try clearing and typing to filter, in case the country isn't initially visible
        try:
            inp = get_country_input(driver)
            inp.clear()
            inp.send_keys(country_name)
            time.sleep(1.5)
            items = driver.find_elements(
                By.CSS_SELECTOR, "div.el-select-dropdown.company-search-country li.el-select-dropdown__item"
            )
            for it in items:
                if safe_text(it).lower() == country_name.lower():
                    target = it
                    break
        except Exception:
            pass
    if target is None:
        logger.warning("country %r not found in dropdown", country_name)
        return False
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target)
    target.click()
    time.sleep(1.5)
    return True


def click_search(driver: WebDriver) -> None:
    btn = driver.find_element(By.CSS_SELECTOR, ".company-search-right-search")
    driver.execute_script("arguments[0].click();", btn)
    time.sleep(8)


# ──────────────────────────────────────────────────────────────────────────
# Company list + pagination
# ──────────────────────────────────────────────────────────────────────────


def collect_company_urls(driver: WebDriver) -> list[str]:
    cards = driver.find_elements(By.CSS_SELECTOR, "ul.membership-list-content-center-list > li")
    urls: list[str] = []
    seen: set[str] = set()
    for card in cards:
        try:
            link = card.find_element(By.CSS_SELECTOR, "a[href*='/en/company/']")
            href = link.get_attribute("href")
            if href and href not in seen:
                seen.add(href)
                urls.append(href)
        except NoSuchElementException:
            continue
    return urls


def click_next_page(driver: WebDriver) -> bool:
    """Click the pagination 'next' button. Return False if there is no next page."""
    next_btns = driver.find_elements(By.CSS_SELECTOR, "button.btn-next")
    for btn in next_btns:
        try:
            if not btn.is_displayed():
                continue
            disabled = btn.get_attribute("disabled")
            cls = btn.get_attribute("class") or ""
            if disabled or "is-disabled" in cls:
                logger.info("next page button disabled — last page reached")
                return False
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(8)
            return True
        except Exception as exc:
            logger.warning("error clicking next page: %s", exc)
            return False
    return False


# ──────────────────────────────────────────────────────────────────────────
# Company detail extraction
# ──────────────────────────────────────────────────────────────────────────


@dataclass
class CompanyInfo:
    country: str
    url: str
    company_name: str = ""
    member_id: str = ""
    year: str = ""
    location: str = ""
    website: str = ""
    main_business: str = ""
    sea_freight: str = ""
    air_freight: str = ""
    contacts: list["ContactInfo"] = field(default_factory=list)


@dataclass
class ContactInfo:
    name: str = ""
    position: str = ""
    email: str = ""
    phone: str = ""
    wechat: str = ""
    whatsapp: str = ""
    skype: str = ""


def is_suspended(driver: WebDriver) -> bool:
    """Detect if the company page indicates it is NOT a JCtrans member."""
    indicators = [
        "It is NOT a JCtrans member",
        "MEMBERSHIP SUSPEND",
    ]
    body_text = ""
    try:
        body_text = driver.find_element(By.TAG_NAME, "main").text
    except NoSuchElementException:
        try:
            body_text = driver.find_element(By.TAG_NAME, "body").text
        except NoSuchElementException:
            pass
    for s in indicators:
        if s.lower() in body_text.lower():
            return True
    return False


def _text_or_empty(driver: WebDriver, css: str | None = None, xpath: str | None = None) -> str:
    try:
        if css:
            el = driver.find_element(By.CSS_SELECTOR, css)
        else:
            el = driver.find_element(By.XPATH, xpath)
        return safe_text(el)
    except NoSuchElementException:
        return ""


def extract_company_name(driver: WebDriver) -> str:
    # The first <p class="text-[18px] font-bold break-word"> in main is the company name
    try:
        els = driver.find_elements(By.XPATH, "//main//p[contains(@class,'font-bold') and contains(@class,'break-word')]")
        for el in els:
            t = safe_text(el)
            if t:
                return t
    except Exception:
        pass
    # Fallback to <h1> or document title
    try:
        return safe_text(driver.find_element(By.TAG_NAME, "h1"))
    except NoSuchElementException:
        return ""


def extract_member_id(driver: WebDriver) -> str:
    """Member ID is in the header near a label 'Member ID'."""
    try:
        el = driver.find_element(
            By.XPATH,
            "//*[contains(translate(normalize-space(.),'MEMBERID','memberid'),'member id')]"
            "/following::b[1]",
        )
        return safe_text(el)
    except NoSuchElementException:
        pass
    # Alternative: any <b> whose preceding sibling text contains 'Member ID'
    try:
        bs = driver.find_elements(By.CSS_SELECTOR, "main b")
        for b in bs:
            parent_text = safe_text(b.find_element(By.XPATH, ".."))
            if "Member ID" in parent_text:
                return safe_text(b)
    except Exception:
        pass
    return ""


def extract_year(driver: WebDriver) -> str:
    try:
        el = driver.find_element(By.XPATH, "//main//span[substring(normalize-space(.), string-length(normalize-space(.)) - 4) = '-Year' or contains(normalize-space(.), '-Year')]")
        return safe_text(el)
    except NoSuchElementException:
        return ""


def extract_desc_box_pair(driver: WebDriver, label: str) -> str:
    """Find a desc-box label paragraph and return text of its following sibling <p>."""
    # Label and value are both <p class="... desc-box">; label has font-bold + text-[#FF6A00].
    try:
        els = driver.find_elements(
            By.XPATH,
            f"//p[contains(@class,'desc-box') and normalize-space(text())='{label}']"
            f"/following-sibling::p[1]",
        )
        for el in els:
            t = safe_text(el)
            if t:
                return t
    except Exception:
        pass
    return ""


def extract_location(driver: WebDriver) -> str:
    # Location label in basic info card
    return extract_desc_box_pair(driver, "Location") or extract_desc_box_pair(driver, "Country/Region")


def extract_website(driver: WebDriver) -> str:
    return extract_desc_box_pair(driver, "Website")


def extract_chips(driver: WebDriver, heading: str) -> str:
    """For a section heading like 'Main Business' return chips joined by ' | '."""
    try:
        heads = driver.find_elements(By.XPATH, f"//main//*[normalize-space(text())='{heading}']")
    except Exception:
        heads = []
    for h in heads:
        try:
            box = h.find_element(By.XPATH, "ancestor::div[contains(@class,'normal-boxShadow')][1]")
        except NoSuchElementException:
            continue
        # chips: leaf divs with bg-[#f5f5f5] and px-4 py-2
        chip_els = box.find_elements(
            By.XPATH,
            ".//div[contains(@class,'px-4') and contains(@class,'py-2') and contains(@class,'rounded')]",
        )
        chips: list[str] = []
        for el in chip_els:
            txt = safe_text(el)
            if not txt:
                continue
            # chips often look like "Air Freight\n34" — take just first line
            first = txt.splitlines()[0].strip()
            if first and first not in chips:
                chips.append(first)
        if chips:
            return " | ".join(chips)
    return ""


def extract_contacts(driver: WebDriver) -> list[ContactInfo]:
    """Iterate .contactCard elements and extract each contact (name/position/icons)."""
    contacts: list[ContactInfo] = []
    cards = driver.find_elements(By.CSS_SELECTOR, "div.contactCard")
    for card in cards:
        try:
            ci = ContactInfo()
            # Name is in .im-info-box (its first text node, possibly with extra Offline/Online tag)
            try:
                name_el = card.find_element(By.CSS_SELECTOR, ".im-info-box")
                # text() of element includes child Online/Offline label too — take first line
                full_name = safe_text(name_el)
                if full_name:
                    ci.name = full_name.splitlines()[0].strip()
            except NoSuchElementException:
                pass
            # Position is the div right after .im-info-box with the gray text style
            try:
                pos_el = card.find_element(
                    By.XPATH, ".//div[contains(@class,'im-info-box')]/following-sibling::div[1]"
                )
                ci.position = safe_text(pos_el)
            except NoSuchElementException:
                pass
            # Chips with icons + content
            chip_els = card.find_elements(
                By.CSS_SELECTOR,
                "div.flex.items-center.justify-start[style*='inline-flex']",
            )
            if not chip_els:
                # fallback: any chip containing .iconView + .content
                chip_els = card.find_elements(
                    By.XPATH, ".//div[.//span[contains(@class,'iconView')] and .//div[contains(@class,'content')]]"
                )
            for chip in chip_els:
                try:
                    icon_html = chip.get_attribute("innerHTML") or ""
                    m = re.search(r'<g\s+id="icon/([\w\-]+)"', icon_html)
                    if not m:
                        continue
                    icon_type = m.group(1).lower()
                    try:
                        content_el = chip.find_element(By.CSS_SELECTOR, "div.content")
                        value = safe_text(content_el)
                    except NoSuchElementException:
                        value = ""
                    if not value:
                        continue
                    if icon_type == "mail":
                        ci.email = value
                    elif icon_type == "phone":
                        ci.phone = value
                    elif icon_type == "wechat":
                        ci.wechat = value
                    elif icon_type == "whatsapp":
                        ci.whatsapp = value
                    elif icon_type == "skype":
                        ci.skype = value
                    # other types (linkedin, facebook, ins) are ignored on purpose
                except Exception as exc:
                    logger.debug("chip parse error: %s", exc)
            if ci.name or ci.email or ci.phone:
                contacts.append(ci)
        except StaleElementReferenceException:
            continue
    return contacts


def parse_company(driver: WebDriver, country: str, url: str) -> CompanyInfo | None:
    """Visit detail page, parse and return CompanyInfo, or None if suspended."""
    safe_get(driver, url, settle=6)
    # Scroll to lazy-load Contact Us section
    for y in (400, 1000, 1800, 2600, 3600, 5000):
        driver.execute_script(f"window.scrollTo(0, {y});")
        time.sleep(0.5)
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(0.6)
    close_blocking_overlays(driver)

    if is_suspended(driver):
        logger.info("skip suspended/non-member: %s", url)
        return None

    info = CompanyInfo(country=country, url=url)
    info.company_name = extract_company_name(driver)
    info.member_id = extract_member_id(driver)
    info.year = extract_year(driver)
    info.location = extract_location(driver)
    info.website = extract_website(driver)
    info.main_business = extract_chips(driver, "Main Business")
    info.sea_freight = extract_chips(driver, "Sea Freight Advantageous")
    info.air_freight = extract_chips(driver, "Air Freight Advantageous")
    info.contacts = extract_contacts(driver)
    return info


# ──────────────────────────────────────────────────────────────────────────
# Excel & checkpoint
# ──────────────────────────────────────────────────────────────────────────


def open_or_create_workbook(path: Path) -> tuple[Workbook, openpyxl.worksheet.worksheet.Worksheet, int]:
    if path.exists():
        wb = load_workbook(path)
        if SHEET_NAME in wb.sheetnames:
            ws = wb[SHEET_NAME]
        else:
            ws = wb.create_sheet(SHEET_NAME)
            ws.append(HEADERS)
        next_row = ws.max_row + 1
        if ws.max_row == 1 and (ws.cell(1, 1).value is None):
            ws.append(HEADERS)
            next_row = 2
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        ws.append(HEADERS)
        next_row = 2

    # Auto-size columns roughly
    widths = [6, 16, 30, 14, 10, 30, 30, 50, 50, 50, 20, 20, 30, 18, 18, 18, 22]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return wb, ws, next_row


def append_company_rows(ws, start_index: int, info: CompanyInfo) -> int:
    """Append one row per contact (or a single row if no contacts). Return next index."""
    contacts = info.contacts or [ContactInfo()]
    idx = start_index
    for c in contacts:
        ws.append([
            idx,                    # A STT
            info.country,           # B
            info.company_name,      # C
            info.member_id,         # D
            info.year,              # E
            info.location,          # F
            info.website,           # G
            info.main_business,     # H
            info.sea_freight,       # I
            info.air_freight,       # J
            c.name,                 # K
            c.position,             # L
            c.email,                # M
            c.phone,                # N
            c.wechat,               # O
            c.whatsapp,             # P
            c.skype,                # Q
        ])
        idx += 1
    return idx


def load_checkpoint() -> dict:
    if CHECKPOINT_PATH.exists():
        try:
            return json.loads(CHECKPOINT_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_checkpoint(state: dict) -> None:
    CHECKPOINT_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


# ──────────────────────────────────────────────────────────────────────────
# Driver flow
# ──────────────────────────────────────────────────────────────────────────


def run_country(driver: WebDriver, country: str, ws, current_index: int,
                save_every: int = 1, save_callback=None,
                max_pages: int | None = None,
                max_companies: int | None = None) -> int:
    """Run scrape for a single country, append to ws, return new running index."""
    logger.info("=== country: %s ===", country)
    safe_get(driver, DIRECTORY_URL, settle=4)
    close_blocking_overlays(driver)
    if not select_country(driver, country):
        logger.warning("could not select country %r — skipping", country)
        return current_index

    click_search(driver)

    page = 1
    companies_done = 0
    while True:
        logger.info("country %s page %s — collecting URLs", country, page)
        urls = collect_company_urls(driver)
        logger.info("  found %s URLs on page %s", len(urls), page)
        if not urls:
            break

        for url in urls:
            if max_companies and companies_done >= max_companies:
                logger.info("max companies hit (%s) — breaking", max_companies)
                return current_index
            logger.info("  -> %s", url)
            try:
                info = parse_company(driver, country, url)
            except Exception as exc:
                logger.error("error parsing %s: %s", url, exc)
                logger.debug(traceback.format_exc())
                continue
            if info is None:
                continue
            current_index = append_company_rows(ws, current_index, info)
            companies_done += 1
            if save_callback and companies_done % save_every == 0:
                save_callback()

        # Re-open the search results — `parse_company` navigated away
        logger.info("returning to result list (page %s)", page)
        safe_get(driver, DIRECTORY_URL, settle=4)
        if not select_country(driver, country):
            logger.warning("country reselect failed; aborting country")
            break
        click_search(driver)
        # Advance to the page we were on
        for _ in range(page - 1):
            if not click_next_page(driver):
                break
        # Now advance to next page
        if max_pages and page >= max_pages:
            logger.info("max_pages hit (%s) — stopping", max_pages)
            break
        if not click_next_page(driver):
            break
        page += 1

    return current_index


def main() -> int:
    parser = argparse.ArgumentParser(description="JCtrans directory scraper")
    parser.add_argument("--country", nargs="*", help="Country name(s) to scrape. Default: all")
    parser.add_argument("--resume", action="store_true", help="Skip countries already in checkpoint")
    parser.add_argument("--max-pages", type=int, default=None, help="Max pages per country (testing)")
    parser.add_argument("--max-companies", type=int, default=None, help="Max companies per country (testing)")
    parser.add_argument("--no-headless", action="store_true", help="Run with visible browser (local only)")
    parser.add_argument("--excel", type=Path, default=EXCEL_PATH, help="Output Excel path")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    setup_logging(args.verbose)
    logger.info("starting jctrans scraper")

    excel_path: Path = args.excel
    wb, ws, next_row = open_or_create_workbook(excel_path)
    state = load_checkpoint()

    driver = make_driver(headless=not args.no_headless)

    def save() -> None:
        wb.save(excel_path)
        save_checkpoint(state)

    try:
        login(driver, LOGIN_EMAIL, LOGIN_PASSWORD)

        if args.country:
            countries = args.country
        else:
            # Fetch fresh list from the dropdown
            safe_get(driver, DIRECTORY_URL, settle=4)
            countries = list_all_countries(driver)
            logger.info("countries to process: %s", len(countries))

        done = set(state.get("completed_countries", []))
        current_index = next_row - 1  # last STT used
        for country in countries:
            if args.resume and country in done:
                logger.info("resume: skip %r (already done)", country)
                continue
            try:
                new_index = run_country(
                    driver, country, ws, current_index + 1,
                    save_every=1,
                    save_callback=save,
                    max_pages=args.max_pages,
                    max_companies=args.max_companies,
                )
                current_index = new_index - 1
            except Exception as exc:
                logger.error("country %s aborted: %s", country, exc)
                logger.debug(traceback.format_exc())
            done.add(country)
            state["completed_countries"] = sorted(done)
            state["last_updated"] = datetime.now().isoformat(timespec="seconds")
            save()
            logger.info("country %s done (%s rows total)", country, ws.max_row - 1)
    finally:
        try:
            wb.save(excel_path)
        except Exception:
            pass
        try:
            save_checkpoint(state)
        except Exception:
            pass
        try:
            driver.quit()
        except Exception:
            pass

    logger.info("done — wrote %s", excel_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
