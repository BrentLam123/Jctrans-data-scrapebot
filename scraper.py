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
import zipfile
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


def make_driver(browser: str = "chrome", headless: bool = True,
                block_images: bool = True, attach: str | None = None) -> WebDriver:
    """Create a WebDriver.

    Args:
        browser: 'chrome' or 'edge'.
        headless: only honoured when launching a fresh browser (ignored on attach).
        block_images: only honoured on a fresh browser launch.
        attach: ``host:port`` of an already-running browser to attach via the
            remote-debugging protocol (e.g. '127.0.0.1:9526'). When given, the
            bot does NOT spawn its own browser — it talks to the one the user
            already opened. ``headless`` / ``block_images`` are ignored in this
            mode because the user owns the launch flags.
    """
    if browser == "edge":
        from selenium.webdriver.edge.options import Options as EdgeOptions
        opts = EdgeOptions()
    else:
        opts = Options()

    if attach:
        opts.add_experimental_option("debuggerAddress", attach)
        # Don't touch other args — user controls the browser
        if browser == "edge":
            driver = webdriver.Edge(options=opts)
        else:
            driver = webdriver.Chrome(options=opts)
        driver.set_page_load_timeout(120)
        return driver

    # Fresh launch
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--lang=en-US")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-background-networking")
    opts.add_argument("--disable-background-timer-throttling")
    opts.add_argument("--disable-renderer-backgrounding")
    opts.add_argument("--disable-backgrounding-occluded-windows")
    opts.add_argument("--disable-notifications")
    if block_images:
        opts.add_argument("--blink-settings=imagesEnabled=false")
    opts.add_argument(f"--user-agent={DEFAULT_USER_AGENT}")
    prefs: dict = {
        "profile.default_content_setting_values.notifications": 2,
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
    }
    if block_images:
        prefs["profile.managed_default_content_settings.images"] = 2
        prefs["profile.default_content_setting_values.plugins"] = 2
    opts.add_experimental_option("prefs", prefs)

    if browser == "edge":
        edge_bin = os.environ.get("EDGE_BIN")
        if edge_bin and Path(edge_bin).exists():
            opts.binary_location = edge_bin
        driver = webdriver.Edge(options=opts)
    else:
        chrome_bin = os.environ.get("CHROME_BIN")
        if chrome_bin and Path(chrome_bin).exists():
            opts.binary_location = chrome_bin
        elif Path("/home/ubuntu/.local/bin/google-chrome").exists():
            opts.binary_location = "/home/ubuntu/.local/bin/google-chrome"
        driver = webdriver.Chrome(options=opts)
    driver.set_page_load_timeout(120)
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


def is_logged_in(driver: WebDriver) -> bool:
    """Detect login state. A visible SIGN IN button means we are logged out."""
    try:
        for b in driver.find_elements(By.CSS_SELECTOR, "button.login-btn"):
            try:
                if b.is_displayed():
                    return False
            except StaleElementReferenceException:
                continue
    except Exception:
        pass
    return True


def ensure_logged_in(driver: WebDriver, email: str, password: str,
                     navigate_first: bool = False) -> bool:
    """If logged out, run login(). Returns True if a re-login actually happened."""
    if navigate_first:
        safe_get(driver, DIRECTORY_URL, settle=2)
    # Dismiss any "logged elsewhere" / blocking popup before checking
    click_confirm_logged_elsewhere(driver)
    close_blocking_overlays(driver)
    if is_logged_in(driver):
        return False
    logger.warning("session lost — re-logging in")
    login(driver, email, password)
    return True


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


# Aliases / ISO 3166-1 alpha-2 codes that we accept as `--country` input and
# map to the canonical jctrans dropdown name. Anything outside this list will
# be matched against the dropdown text directly (case-insensitive).
COUNTRY_ALIASES: dict[str, str] = {
    # ISO-2 codes
    "gb": "United Kingdom",
    "uk": "United Kingdom",
    "us": "United States",
    "usa": "United States",
    "jp": "Japan",
    "vn": "Vietnam",
    "cn": "China",
    "hk": "Hong Kong, China",
    "tw": "Taiwan-China",
    "kr": "South Korea",
    "kp": "North Korea",
    "sg": "Singapore",
    "my": "Malaysia",
    "id": "Indonesia",
    "ph": "Philippines",
    "th": "Thailand",
    "in": "India",
    "pk": "Pakistan",
    "bd": "Bangladesh",
    "lk": "Sri Lanka",
    "np": "Nepal",
    "au": "Australia",
    "nz": "New Zealand",
    "de": "Germany",
    "fr": "France",
    "it": "Italy",
    "es": "Spain",
    "pt": "Portugal",
    "nl": "Netherlands",
    "be": "Belgium",
    "lu": "Luxembourg",
    "ch": "Switzerland",
    "at": "Austria",
    "ie": "Ireland",
    "se": "Sweden",
    "no": "Norway",
    "dk": "Denmark",
    "fi": "Finland",
    "is": "Iceland",
    "pl": "Poland",
    "cz": "Czechia",
    "sk": "Slovakia",
    "hu": "Hungary",
    "ro": "Romania",
    "bg": "Bulgaria",
    "gr": "Greece",
    "tr": "Turkiye",
    "ru": "Russia",
    "ua": "Ukraine",
    "by": "Belarus",
    "ee": "Estonia",
    "lv": "Latvia",
    "lt": "Lithuania",
    "rs": "Serbia",
    "hr": "Croatia",
    "si": "Slovenia",
    "ba": "Bosnia and Herzegovina",
    "mk": "North Macedonia",
    "al": "Albania",
    "ca": "Canada",
    "mx": "Mexico",
    "br": "Brazil",
    "ar": "Argentina",
    "cl": "Chile",
    "co": "Colombia",
    "pe": "Peru",
    "ve": "Venezuela",
    "uy": "Uruguay",
    "ec": "Ecuador",
    "bo": "Bolivia",
    "py": "Paraguay",
    "ae": "United Arab Emirates",
    "sa": "Saudi Arabia",
    "qa": "Qatar",
    "kw": "Kuwait",
    "om": "Oman",
    "bh": "Bahrain",
    "jo": "Jordan",
    "lb": "Lebanon",
    "il": "Israel",
    "ir": "Iran",
    "iq": "Iraq",
    "eg": "Egypt",
    "ma": "Morocco",
    "tn": "Tunisia",
    "dz": "Algeria",
    "ly": "Libya",
    "ng": "Nigeria",
    "za": "South Africa",
    "ke": "Kenya",
    "et": "Ethiopia",
    "gh": "Ghana",
    "ci": "Cote D'Ivoire",
    "sn": "Senegal",
    "ug": "Uganda",
    "tz": "Tanzania",
    # Common spelling variants jctrans uses non-standard forms for
    "great britain": "United Kingdom",
    "england": "United Kingdom",
    "scotland": "United Kingdom",
    "wales": "United Kingdom",
    "united states of america": "United States",
    "america": "United States",
    "korea": "South Korea",
    "korea, republic of": "South Korea",
    "republic of korea": "South Korea",
    "korea, south": "South Korea",
    "korea, north": "North Korea",
    "russian federation": "Russia",
    "viet nam": "Vietnam",
    "hong kong": "Hong Kong, China",
    "taiwan": "Taiwan-China",
    "turkey": "Turkiye",
    "ivory coast": "Cote D'Ivoire",
    "czech republic": "Czechia",
}


# Junk entries that show up in the jctrans dropdown but are NOT real
# countries (e.g. orphan 2-letter codes that return mixed/wrong data).
# Anything we see here will be filtered out of list_all_countries().
_KNOWN_JUNK_COUNTRY_ENTRIES = {"gb", "aq"}


def _looks_like_real_country(name: str) -> bool:
    """Heuristic for filtering jctrans dropdown junk entries.

    Real countries:
      - have at least 4 characters (catches all real short names like
        Guam, Niue, Oman, Chad, Peru, Laos, Fiji, Togo)
      - contain at least one space, OR are not all uppercase
        (eliminates ISO-style codes like "GB" / "AQ")
      - are not explicitly blacklisted.
    """
    s = (name or "").strip()
    if len(s) < 4:
        return False
    if s.lower() in _KNOWN_JUNK_COUNTRY_ENTRIES:
        return False
    # A 2-3 letter all-uppercase token (with no spaces) is almost certainly
    # a code, not a real country name.
    if len(s) <= 3 and s.isupper() and " " not in s:
        return False
    return True


def list_all_countries(driver: WebDriver) -> list[str]:
    items = open_country_dropdown(driver)
    names: list[str] = []
    for it in items:
        try:
            if it.is_displayed():
                t = safe_text(it)
                if t and _looks_like_real_country(t):
                    names.append(t)
        except Exception:
            pass
    # close dropdown by pressing escape
    try:
        get_country_input(driver).send_keys(Keys.ESCAPE)
    except Exception:
        pass
    time.sleep(0.5)
    # De-duplicate while preserving order
    seen: set[str] = set()
    out: list[str] = []
    for n in names:
        if n not in seen:
            seen.add(n)
            out.append(n)
    return out


def _norm_country(s: str) -> str:
    """Lowercase + collapse whitespace + strip — for fuzzy country matching."""
    return " ".join((s or "").split()).strip().lower()


def resolve_country_alias(country_name: str) -> str:
    """Translate ISO-2 codes / common variants to the canonical jctrans name."""
    key = _norm_country(country_name)
    return COUNTRY_ALIASES.get(key, country_name)


def select_country(driver: WebDriver, country_name: str) -> bool:
    """Select a country in the jctrans dropdown — strict, never silent-wrong.

    Returns True only if we successfully clicked an option whose normalized
    text equals (or contains) the resolved name. Returns False — and the
    caller MUST skip the country — if no clean match is found.
    """
    # Step 0: resolve aliases / ISO codes up-front, so "GB" → "United Kingdom".
    resolved = resolve_country_alias(country_name)
    if resolved != country_name:
        logger.info("country alias %r → %r", country_name, resolved)
    want = _norm_country(resolved)
    if not want:
        logger.warning("empty country name — refusing to select")
        return False
    if not _looks_like_real_country(resolved):
        logger.warning(
            "country %r looks like a junk dropdown entry — refusing to select",
            country_name,
        )
        return False

    items = open_country_dropdown(driver)
    target = None
    # Pass 1: exact match against currently visible options
    for it in items:
        if _norm_country(safe_text(it)) == want:
            target = it
            break
    # Pass 2: type into the filter input (el-select filterable), exact match
    if target is None:
        try:
            inp = get_country_input(driver)
            inp.clear()
            inp.send_keys(resolved)
            time.sleep(1.5)
            items = driver.find_elements(
                By.CSS_SELECTOR,
                "div.el-select-dropdown.company-search-country li.el-select-dropdown__item",
            )
            for it in items:
                if _norm_country(safe_text(it)) == want:
                    target = it
                    break
        except Exception:
            pass
    # Pass 3: STRICT prefix match — only when the filter narrows to a small
    # set AND the option starts with the requested name (case-insensitive).
    # Prevents the old "GB → first remaining unrelated option" silent
    # mismatch by refusing to pick anything that doesn't actually start with
    # the requested string.
    if target is None:
        try:
            visible = [
                it for it in driver.find_elements(
                    By.CSS_SELECTOR,
                    "div.el-select-dropdown.company-search-country li.el-select-dropdown__item",
                )
                if it.is_displayed() and safe_text(it)
            ]
            for it in visible:
                norm = _norm_country(safe_text(it))
                if norm.startswith(want + " ") or norm.startswith(want + ","):
                    target = it
                    logger.info(
                        "country %r matched as %r (prefix fallback)",
                        country_name, safe_text(it),
                    )
                    break
        except Exception:
            pass
    if target is None:
        logger.warning(
            "country %r (resolved=%r) not found in dropdown — skipping",
            country_name, resolved,
        )
        # Close dropdown so we don't leave it open for the next country
        try:
            get_country_input(driver).send_keys(Keys.ESCAPE)
        except Exception:
            pass
        return False
    # Sanity-check: verify the option's text actually matches before clicking.
    picked = _norm_country(safe_text(target))
    if picked != want and not (
        picked.startswith(want + " ") or picked.startswith(want + ",")
    ):
        logger.warning(
            "country %r matched %r — looks wrong, refusing to click",
            country_name, safe_text(target),
        )
        try:
            get_country_input(driver).send_keys(Keys.ESCAPE)
        except Exception:
            pass
        return False
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target)
    target.click()
    time.sleep(1.5)
    return True


def _wait_for_result_list(driver: WebDriver, timeout: int = 20) -> bool:
    """Wait until at least one company card (or 'no results') is rendered."""
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: bool(d.find_elements(
                By.CSS_SELECTOR,
                "ul.membership-list-content-center-list > li a[href*='/en/company/']",
            )) or bool(d.find_elements(By.XPATH, "//*[contains(.,'No Data') or contains(.,'no data')]"))
        )
        return True
    except TimeoutException:
        return False


def click_search(driver: WebDriver) -> None:
    btn = driver.find_element(By.CSS_SELECTOR, ".company-search-right-search")
    driver.execute_script("arguments[0].click();", btn)
    _wait_for_result_list(driver, timeout=25)


# ──────────────────────────────────────────────────────────────────────────
# Company list + pagination
# ──────────────────────────────────────────────────────────────────────────


def _collect_company_urls_once(driver: WebDriver) -> list[str]:
    """Return the *currently rendered* company URLs in one CDP round-trip.

    No retry — used both directly (during settle polling) and by the
    retry wrapper below.
    """
    try:
        result = driver.execute_script(
            "const out = [];"
            "const seen = new Set();"
            "const links = document.querySelectorAll("
            "  'ul.membership-list-content-center-list > li a[href*=\"/en/company/\"]'"
            ");"
            "for (const a of links) {"
            "  const h = a.href;"
            "  if (h && !seen.has(h)) { seen.add(h); out.push(h); }"
            "}"
            "return out;"
        )
        return list(result or [])
    except Exception as exc:
        logger.warning("collect_company_urls JS path failed: %s — falling back", exc)
    # Fallback: WebElement-based, with stale guards.
    urls: list[str] = []
    seen: set[str] = set()
    try:
        cards = driver.find_elements(
            By.CSS_SELECTOR, "ul.membership-list-content-center-list > li"
        )
    except Exception:
        return urls
    for card in cards:
        try:
            link = card.find_element(By.CSS_SELECTOR, "a[href*='/en/company/']")
            href = link.get_attribute("href")
            if href and href not in seen:
                seen.add(href)
                urls.append(href)
        except (NoSuchElementException, StaleElementReferenceException):
            continue
    return urls


def collect_company_urls(driver: WebDriver, max_wait: float = 12.0,
                          poll: float = 0.5) -> list[str]:
    """Collect company detail URLs from the current results page.

    Vue can leave the list momentarily empty between pagination clicks
    (clear → API call → re-render). We poll up to ``max_wait`` seconds for
    a non-empty, *stable* card list. ``stable`` here means: two consecutive
    polls with identical hrefs. If we never see a stable non-empty result,
    we return whatever the last poll produced (possibly empty).
    """
    deadline = time.time() + max_wait
    prev: list[str] = []
    last_non_empty: list[str] = []
    while True:
        urls = _collect_company_urls_once(driver)
        if urls:
            last_non_empty = urls
            # Two identical, non-empty polls in a row → settled.
            if prev and prev == urls:
                return urls
        prev = urls
        if time.time() >= deadline:
            return last_non_empty
        time.sleep(poll)


def _current_page_num(driver: WebDriver) -> str:
    """Return the active page number text from el-pagination, or '' if not found."""
    for sel in (
        "ul.el-pager li.is-active",
        "ul.el-pager li.active",
        ".el-pagination li.number.active",
    ):
        try:
            for el in driver.find_elements(By.CSS_SELECTOR, sel):
                t = (el.text or "").strip()
                if t:
                    return t
        except Exception:
            pass
    return ""


def _first_card_href(driver: WebDriver) -> str:
    """Return the href of the first company card link, or ''.

    Uses a JS lookup so it never raises StaleElementReferenceException.
    """
    try:
        result = driver.execute_script(
            "const a = document.querySelector("
            "  'ul.membership-list-content-center-list > li a[href*=\"/en/company/\"]'"
            ");"
            "return a ? a.href : '';"
        )
        return result or ""
    except Exception:
        return ""


def _all_card_hrefs(driver: WebDriver) -> list[str]:
    """Return *all* current card hrefs (used to detect real list change)."""
    try:
        result = driver.execute_script(
            "return Array.from(document.querySelectorAll("
            "  'ul.membership-list-content-center-list > li a[href*=\"/en/company/\"]'"
            ")).map(a => a.href);"
        )
        return list(result or [])
    except Exception:
        return []


def _find_enabled_next_btn(driver: WebDriver) -> WebElement | None:
    """Return a fresh, enabled .btn-next element (or None)."""
    for b in driver.find_elements(By.CSS_SELECTOR, "button.btn-next"):
        try:
            if b.get_attribute("disabled"):
                continue
            cls = b.get_attribute("class") or ""
            if "is-disabled" in cls:
                continue
            return b
        except StaleElementReferenceException:
            continue
    return None


def _next_btn_state(driver: WebDriver) -> str:
    """Return 'enabled' | 'disabled' | 'missing' — uses fresh lookup each call."""
    btns = driver.find_elements(By.CSS_SELECTOR, "button.btn-next")
    if not btns:
        return "missing"
    enabled = False
    for b in btns:
        try:
            cls = b.get_attribute("class") or ""
            disabled = b.get_attribute("disabled")
            if not disabled and "is-disabled" not in cls:
                enabled = True
                break
        except StaleElementReferenceException:
            continue
    return "enabled" if enabled else "disabled"


def _page_num_int(driver: WebDriver) -> int | None:
    try:
        return int(_current_page_num(driver) or "")
    except (ValueError, TypeError):
        return None


def _click_next_page_inner(driver: WebDriver) -> bool:
    # Make sure pagination is in viewport.
    try:
        driver.execute_script(
            "const b = document.querySelector('button.btn-next');"
            "if (b) b.scrollIntoView({block:'center'});"
        )
    except Exception:
        pass

    # Wait for at least one .btn-next to be present (Vue may re-render briefly).
    try:
        WebDriverWait(driver, 10).until(
            lambda d: bool(d.find_elements(By.CSS_SELECTOR, "button.btn-next"))
        )
    except TimeoutException:
        logger.info("no .btn-next found in DOM — assuming single-page result")
        return False

    state = _next_btn_state(driver)
    if state == "missing":
        return False
    if state == "disabled":
        logger.info("next page button disabled — last page reached")
        return False

    num_before = _page_num_int(driver)
    hrefs_before = set(_all_card_hrefs(driver))
    logger.info(
        "pagination: clicking next (current page=%s, cards=%s)",
        num_before, len(hrefs_before),
    )

    # Closure used by `_advanced` to require the post-click card set to be
    # *stable* across two consecutive polls. Without this we can latch onto
    # a transient state (cards present briefly, then cleared by Vue while
    # the API response is still in flight) and declare advance success
    # while the list is actually empty, causing the loop to think the
    # country has no more pages.
    last_seen_set: list[set[str]] = [set()]

    def _advanced(d: WebDriver) -> bool:
        """Return True only when:
          - the page number advanced (if we know it),
          - the card list contains at least one valid /en/company/ href that
            isn't part of the previous page,
          - *and* the same card set has been observed across two
            consecutive polls (i.e. Vue has finished re-rendering).
        """
        try:
            new_num = _page_num_int(d)
            new_hrefs = _all_card_hrefs(d)
            # Cards must be present and all hrefs must be non-empty.
            if not new_hrefs or any(not h for h in new_hrefs):
                last_seen_set[0] = set()
                return False
            new_set = set(new_hrefs)
            # Page number must have advanced (when both numbers are known).
            if num_before is not None and new_num is not None:
                if new_num <= num_before:
                    last_seen_set[0] = set()
                    return False
            # Card list must actually move forward — at least one fresh
            # href that wasn't on the previous page.
            if hrefs_before and new_set == hrefs_before:
                last_seen_set[0] = set()
                return False
            if hrefs_before and new_set.issubset(hrefs_before):
                last_seen_set[0] = set()
                return False
            # Settle check: the same set must show up on TWO consecutive
            # polls (~500ms apart). Otherwise Vue is still re-rendering.
            if last_seen_set[0] != new_set:
                last_seen_set[0] = new_set
                return False
            return True
        except StaleElementReferenceException:
            last_seen_set[0] = set()
            return False

    # Click — JS click + always re-find the button (avoid holding a stale ref)
    for attempt in (1, 2):
        btn = _find_enabled_next_btn(driver)
        if btn is None:
            logger.info("next-page button vanished before click — assume done")
            return False
        try:
            driver.execute_script("arguments[0].click();", btn)
        except StaleElementReferenceException:
            # Re-find and try once more
            btn = _find_enabled_next_btn(driver)
            if btn is None:
                return False
            try:
                driver.execute_script("arguments[0].click();", btn)
            except Exception as exc:
                logger.warning("retry click failed: %s", exc)
        try:
            WebDriverWait(driver, 35).until(_advanced)
        except TimeoutException:
            if attempt == 1:
                logger.warning(
                    "next-page click did not produce a change in 35s — retrying"
                )
                continue
            logger.error(
                "next-page transition still failing after retry — giving up on pagination"
            )
            return False
        # Got an advance — wait for the new card list to render.
        try:
            WebDriverWait(driver, 15).until(
                lambda d: bool(d.find_elements(
                    By.CSS_SELECTOR,
                    "ul.membership-list-content-center-list > li a[href*='/en/company/']",
                ))
            )
        except TimeoutException:
            logger.warning("page transition detected but cards never re-rendered")
        num_after = _page_num_int(driver)
        href_after = _first_card_href(driver)
        logger.info(
            "pagination: advanced to page=%s, first href=%s",
            num_after, href_after.rsplit("/", 1)[-1] if href_after else None,
        )
        return True
    return False


def click_next_page(driver: WebDriver) -> bool:
    """Click the pagination 'next' button.

    Always returns a bool — never raises. False means 'no more pages' / 'gave
    up'. True means 'now on a new page with cards rendered'.
    """
    try:
        return _click_next_page_inner(driver)
    except Exception as exc:
        logger.error("click_next_page crashed: %s", exc)
        logger.debug(traceback.format_exc())
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


# Sentinel returned by _parse_current_tab when the session got kicked.
SESSION_LOST = object()

# Single-pass JS extractor — pulls every field we want in ONE CDP round-trip.
# This is the biggest perf lever on a remote (Windows) machine where each
# Selenium command costs 50–100 ms over the wire — the old per-field
# extractors did ~30–50 round-trips per detail tab, ≈3 s per company in
# pure overhead. With this one call we cut the per-company budget by
# roughly a factor of 2–3.
_EXTRACT_JS = r"""
return (function() {
    const out = {
        company_name: '', member_id: '', year: '', location: '',
        website: '', main_business: '', sea_freight: '', air_freight: '',
        contacts: [], is_suspended: false, logged_out: false, ready: false,
    };
    // Logged-out: visible "SIGN IN" button => server is rendering masked data
    const loginBtn = Array.from(document.querySelectorAll('button.login-btn'))
        .find(b => b.offsetParent !== null);
    if (loginBtn) { out.logged_out = true; return out; }

    const mainEl = document.querySelector('main') || document.body;
    const mainText = (mainEl.innerText || '').toLowerCase();
    if (mainText.includes('it is not a jctrans member')
        || mainText.includes('membership suspend')) {
        out.is_suspended = true;
        return out;
    }

    // Company name: first <p> with both font-bold + break-word in main
    const nameEls = mainEl.querySelectorAll(
        'p[class*="font-bold"][class*="break-word"]'
    );
    for (const el of nameEls) {
        const t = (el.innerText || '').trim();
        if (t) { out.company_name = t; break; }
    }
    if (!out.company_name) {
        const h1 = document.querySelector('h1');
        if (h1) out.company_name = (h1.innerText || '').trim();
    }
    if (!out.company_name) { return out; }
    out.ready = true;

    // Member ID — find an element whose text contains "Member ID" and
    // grab the next <b> in document order.
    (function() {
        const all = mainEl.querySelectorAll('*');
        for (const el of all) {
            if (el.children.length > 8) continue;
            const t = (el.textContent || '').trim().toLowerCase();
            if (!t.includes('member id')) continue;
            const walker = document.createTreeWalker(
                mainEl, NodeFilter.SHOW_ELEMENT, null
            );
            walker.currentNode = el;
            let n = walker.nextNode();
            let safety = 60;
            while (n && safety-- > 0) {
                if (n.tagName === 'B') {
                    const bt = (n.innerText || '').trim();
                    if (bt && !bt.toLowerCase().includes('member id')) {
                        out.member_id = bt; return;
                    }
                }
                n = walker.nextNode();
            }
            return;
        }
    })();

    // Year (e.g. "15-Year")
    {
        const re = /(\d+\s*-\s*Year)/i;
        const spans = mainEl.querySelectorAll('span');
        for (const s of spans) {
            const m = (s.innerText || '').match(re);
            if (m) { out.year = m[1]; break; }
        }
    }

    // desc-box label/value pair (Location, Website, ...)
    function descBoxPair(label) {
        const ps = mainEl.querySelectorAll('p[class*="desc-box"]');
        for (const p of ps) {
            if ((p.innerText || '').trim() === label) {
                const sib = p.nextElementSibling;
                if (sib && sib.tagName === 'P') {
                    return (sib.innerText || '').trim();
                }
            }
        }
        return '';
    }
    out.location = descBoxPair('Location') || descBoxPair('Country/Region');
    out.website = descBoxPair('Website');

    // Chips block: heading element → ancestor with .normal-boxShadow → chips
    function chipsForHeading(heading) {
        const all = mainEl.querySelectorAll('*');
        for (const h of all) {
            if (h.children.length !== 0) continue;
            if ((h.innerText || '').trim() !== heading) continue;
            let box = h;
            while (box) {
                const cls = (typeof box.className === 'string')
                    ? box.className : '';
                if (cls.includes('normal-boxShadow')) break;
                box = box.parentElement;
            }
            if (!box) continue;
            const chipEls = box.querySelectorAll(
                'div[class*="px-4"][class*="py-2"][class*="rounded"]'
            );
            const chips = []; const seen = new Set();
            for (const c of chipEls) {
                const txt = (c.innerText || '').trim();
                if (!txt) continue;
                const first = txt.split('\n')[0].trim();
                if (first && !seen.has(first)) {
                    seen.add(first); chips.push(first);
                }
            }
            if (chips.length) return chips.join(' | ');
        }
        return '';
    }
    out.main_business = chipsForHeading('Main Business');
    out.sea_freight = chipsForHeading('Sea Freight Advantageous');
    out.air_freight = chipsForHeading('Air Freight Advantageous');

    // Contacts — one entry per .contactCard
    const cards = document.querySelectorAll('div.contactCard');
    for (const card of cards) {
        const ci = {
            name: '', position: '', email: '', phone: '',
            wechat: '', whatsapp: '', skype: '',
        };
        const nameEl = card.querySelector('.im-info-box');
        if (nameEl) {
            const t = (nameEl.innerText || '').trim();
            if (t) ci.name = t.split('\n')[0].trim();
            // Position = first sibling DIV after .im-info-box
            let sib = nameEl.nextElementSibling;
            while (sib && sib.tagName !== 'DIV') sib = sib.nextElementSibling;
            if (sib) ci.position = (sib.innerText || '').trim();
        }
        // Chips: slot-based detection. The contact card has a horizontal
        // chip container whose direct children are slot wrappers, in this
        // fixed order: [mail, phone, wechat, whatsapp, skype, ...socials].
        // Empty slots render as <div><!----></div>. mail/phone use an
        // <g id="icon/X"> SVG we can sniff directly, but whatsapp/skype use
        // a different iconfont SVG with no identifier — so we map by slot
        // index as a fallback.
        const SLOT_ORDER = ['mail', 'phone', 'wechat', 'whatsapp', 'skype'];
        let chipContainer = null;
        const firstChip = card.querySelector(
            'div.flex.items-center.justify-start[style*="inline-flex"]'
        );
        if (firstChip && firstChip.parentElement) {
            chipContainer = firstChip.parentElement.parentElement;
        }
        if (!chipContainer) {
            chipContainer = card.querySelector(
                'div.flex.flex-wrap, div[class*="flex-wrap"]'
            );
        }
        if (chipContainer) {
            const slots = Array.from(chipContainer.children);
            for (let i = 0; i < slots.length; i++) {
                const slot = slots[i];
                const chip = slot.querySelector(
                    'div.flex.items-center.justify-start, div.flex.items-center'
                );
                if (!chip) continue;
                const contentEl = chip.querySelector('div.content');
                const value = contentEl ? (contentEl.innerText || '').trim() : '';
                if (!value) continue;
                // Prefer the icon id if present, fall back to slot position.
                const html = chip.innerHTML || '';
                const m = html.match(/<g\s+id="icon\/([\w\-]+)"/i);
                let iconType = m ? m[1].toLowerCase() : '';
                if (!iconType && i < SLOT_ORDER.length) {
                    iconType = SLOT_ORDER[i];
                }
                if (iconType === 'mail') ci.email = value;
                else if (iconType === 'phone') ci.phone = value;
                else if (iconType === 'wechat') ci.wechat = value;
                else if (iconType === 'whatsapp') ci.whatsapp = value;
                else if (iconType === 'skype') ci.skype = value;
            }
        }
        if (ci.name || ci.email || ci.phone) out.contacts.push(ci);
    }
    return out;
})();
"""


def _extract_all_via_js(driver: WebDriver) -> dict | None:
    try:
        return driver.execute_script(_EXTRACT_JS)
    except Exception as exc:
        logger.debug("extract JS failed: %s", exc)
        return None


def _parse_current_tab(driver: WebDriver, country: str, url: str,
                       wait_secs: int = 25):
    """Parse the detail page that is already loaded in the current tab.

    Returns:
        CompanyInfo — full data extracted.
        None        — company is suspended/non-member; skip.
        SESSION_LOST — the page rendered as logged-out; caller should retry.
    """
    # Kick lazy-load (IntersectionObserver) before we start polling.
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    except Exception:
        pass

    deadline = time.time() + wait_secs
    data: dict | None = None
    poll = 0.15
    while time.time() < deadline:
        data = _extract_all_via_js(driver)
        if data is None:
            time.sleep(poll)
            continue
        if data.get("logged_out"):
            logger.warning("logged out (detected on %s)", url)
            return SESSION_LOST
        if data.get("is_suspended"):
            logger.info("skip suspended/non-member: %s", url)
            return None
        if data.get("ready"):
            break
        time.sleep(poll)

    if not data or not data.get("ready"):
        logger.warning("could not parse %s (slow / never finished)", url)
        return None

    # company_name is up — but contactCard lazy-loads via IntersectionObserver
    # and may take a tick after we scrolled. Quick second-pass for contacts.
    if not data.get("contacts"):
        for _ in range(10):  # up to ~1.5 s extra
            time.sleep(0.15)
            data2 = _extract_all_via_js(driver)
            if data2 and data2.get("contacts"):
                data = data2
                break
            if data2 and data2.get("ready"):
                # Page is ready, still no contacts — keep waiting in case
                # IntersectionObserver hasn't fired yet, but stop after the loop.
                data = data2

    info = CompanyInfo(country=country, url=url)
    info.company_name = data.get("company_name") or ""
    info.member_id = data.get("member_id") or ""
    info.year = data.get("year") or ""
    info.location = data.get("location") or ""
    info.website = data.get("website") or ""
    info.main_business = data.get("main_business") or ""
    info.sea_freight = data.get("sea_freight") or ""
    info.air_freight = data.get("air_freight") or ""
    info.contacts = []
    for c in data.get("contacts") or []:
        info.contacts.append(ContactInfo(
            name=c.get("name") or "",
            position=c.get("position") or "",
            email=c.get("email") or "",
            phone=c.get("phone") or "",
            wechat=c.get("wechat") or "",
            whatsapp=c.get("whatsapp") or "",
            skype=c.get("skype") or "",
        ))
    return info


def parse_company(driver: WebDriver, country: str, url: str) -> CompanyInfo | None:
    """Visit detail page in the current tab and parse it."""
    safe_get(driver, url, settle=1)
    return _parse_current_tab(driver, country, url)


# ──────────────────────────────────────────────────────────────────────────
# Excel & checkpoint
# ──────────────────────────────────────────────────────────────────────────


def open_or_create_workbook(path: Path) -> tuple[Workbook, openpyxl.worksheet.worksheet.Worksheet, int]:
    """Open ``path`` if it's a valid xlsx; otherwise back it up and create a new one."""
    wb: Workbook | None = None
    if path.exists() and path.stat().st_size > 0:
        try:
            wb = load_workbook(path)
        except (zipfile.BadZipFile, KeyError, OSError) as exc:
            backup = path.with_suffix(path.suffix + f".broken-{int(time.time())}.bak")
            try:
                path.rename(backup)
                logger.warning("existing %s is not a valid xlsx (%s) — backed up to %s", path, exc, backup)
            except OSError:
                logger.warning("existing %s is not a valid xlsx (%s) — overwriting", path, exc)
            wb = None

    if wb is None:
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        ws.append(HEADERS)
        next_row = 2
    else:
        if SHEET_NAME in wb.sheetnames:
            ws = wb[SHEET_NAME]
        else:
            ws = wb.create_sheet(SHEET_NAME)
            ws.append(HEADERS)
        next_row = ws.max_row + 1
        if ws.max_row == 1 and (ws.cell(1, 1).value is None):
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


def collect_all_urls_for_country(driver: WebDriver, country: str,
                                 max_pages: int | None = None,
                                 max_urls: int | None = None) -> list[str]:
    """Open the directory, select country, paginate through and collect all URLs.

    Returns a de-duplicated list of /en/company/... URLs in their original order.
    """
    safe_get(driver, DIRECTORY_URL, settle=2)
    close_blocking_overlays(driver)
    ensure_logged_in(driver, LOGIN_EMAIL, LOGIN_PASSWORD)
    if not select_country(driver, country):
        logger.warning("could not select country %r — skipping", country)
        return []
    click_search(driver)

    all_urls: list[str] = []
    seen: set[str] = set()
    page = 1
    while True:
        page_urls = collect_company_urls(driver)
        new = [u for u in page_urls if u not in seen]
        for u in new:
            seen.add(u)
        all_urls.extend(new)
        logger.info("  page %s: %s urls (cumulative %s)", page, len(new), len(all_urls))
        if max_urls and len(all_urls) >= max_urls:
            logger.info("max_urls hit (%s) — stop collecting", max_urls)
            break
        if max_pages and page >= max_pages:
            logger.info("max_pages hit (%s) — stop collecting", max_pages)
            break
        if not click_next_page(driver):
            break
        page += 1
    return all_urls


def _close_extra_tabs(driver: WebDriver, keep_handle: str) -> None:
    """Close every tab except ``keep_handle`` and switch back to it."""
    for h in list(driver.window_handles):
        if h == keep_handle:
            continue
        try:
            driver.switch_to.window(h)
            driver.close()
        except Exception:
            pass
    try:
        driver.switch_to.window(keep_handle)
    except Exception:
        pass


def _open_urls_in_tabs(driver: WebDriver, urls: list[str], stagger: float = 0.05) -> list[str]:
    """Open each URL in a new background tab, return the list of new handles in order.

    Uses W3C ``driver.switch_to.new_window('tab')`` instead of ``window.open()``
    because:

    * It is a WebDriver-level command (CDP), so it bypasses the browser's
      popup blocker — which is critical when *attached* to a real Edge or
      Chrome window where ``window.open()`` from a non-user-gesture context
      is silently blocked.
    * It is reliable when 20 tabs are opened in rapid succession.

    The tab is navigated using ``location.replace(...)`` (non-blocking) so
    detail pages load *in parallel* while the caller continues parsing
    earlier tabs.
    """
    main_handle = driver.current_window_handle
    new_handles: list[str] = []
    before = set(driver.window_handles)
    for url in urls:
        try:
            driver.switch_to.new_window("tab")
        except Exception as exc:
            # Fallback to JS open if CDP for some reason isn't available
            logger.warning("switch_to.new_window failed: %s — falling back to window.open", exc)
            try:
                driver.switch_to.window(main_handle)
            except Exception:
                pass
            try:
                driver.execute_script("window.open(arguments[0], '_blank');", url)
            except Exception:
                continue
            # Find the newly-opened handle
            for h in driver.window_handles:
                if h not in before and h not in new_handles:
                    new_handles.append(h)
                    break
            if stagger:
                time.sleep(stagger)
            continue

        h = driver.current_window_handle
        new_handles.append(h)
        # Navigate without blocking selenium — let the page load in the
        # background while we keep opening more tabs.
        try:
            driver.execute_script(
                "window.location.replace(arguments[0]);", url
            )
        except Exception:
            try:
                driver.get(url)
            except Exception:
                pass
        if stagger:
            time.sleep(stagger)
    # Always end on the main results tab
    try:
        driver.switch_to.window(main_handle)
    except Exception:
        pass
    if len(new_handles) < len(urls):
        logger.warning("opened only %s/%s tabs", len(new_handles), len(urls))
    return new_handles


def _open_country_results(driver: WebDriver, country: str,
                          email: str, password: str) -> bool:
    """Navigate to the directory, ensure logged-in, select country and click search."""
    safe_get(driver, DIRECTORY_URL, settle=2)
    close_blocking_overlays(driver)
    ensure_logged_in(driver, email, password, navigate_first=False)
    if not select_country(driver, country):
        logger.warning("could not select country %r — skipping", country)
        return False
    click_search(driver)
    return True


def _scrape_page_tabs(driver: WebDriver, country: str, urls: list[str],
                      main_handle: str,
                      tab_batch: int = 20) -> tuple[list[CompanyInfo], bool]:
    """Open ``urls`` as tabs (in batches of ``tab_batch``), scrape and close.

    Returns ``(results, session_lost)``. If ``session_lost`` is True, the caller
    must throw away the partial results, re-login and retry the page.
    """
    results: list[CompanyInfo] = []
    session_lost = False
    for start in range(0, len(urls), tab_batch):
        batch = urls[start:start + tab_batch]
        logger.info("opening %s tabs (%s..%s of %s)",
                    len(batch), start + 1, start + len(batch), len(urls))
        new_handles = _open_urls_in_tabs(driver, batch)
        for h, url in zip(new_handles, batch):
            try:
                driver.switch_to.window(h)
            except Exception as exc:
                logger.warning("could not switch to tab for %s: %s", url, exc)
                continue
            try:
                info = _parse_current_tab(driver, country, url)
            except Exception as exc:
                logger.error("error parsing %s: %s", url, exc)
                logger.debug(traceback.format_exc())
                info = None
            finally:
                try:
                    driver.close()
                except Exception:
                    pass
            if info is SESSION_LOST:
                logger.warning("session lost on %s — aborting batch", url)
                session_lost = True
                continue
            if info is not None:
                results.append(info)
        # Always end the batch on the results tab
        try:
            driver.switch_to.window(main_handle)
        except Exception:
            logger.error("results tab vanished while scraping batch — aborting page")
            session_lost = True
            break
        if session_lost:
            break
    return results, session_lost


def run_country(driver: WebDriver, country: str, ws, current_index: int,
                save_every: int = 1, save_callback=None,
                max_pages: int | None = None,
                max_companies: int | None = None,
                use_tabs: bool = True,
                tab_batch: int = 20) -> int:
    """Run scrape for a single country, append to ws, return new running index."""
    logger.info("=== country: %s ===", country)
    if not use_tabs:
        return _run_country_serial(
            driver, country, ws, current_index,
            save_every=save_every, save_callback=save_callback,
            max_pages=max_pages, max_companies=max_companies,
        )

    if not _open_country_results(driver, country, LOGIN_EMAIL, LOGIN_PASSWORD):
        return current_index

    main_handle = driver.current_window_handle
    seen_urls: set[str] = set()  # detail urls already written to Excel
    companies_done = 0
    started = time.time()
    page = 1
    retries_this_page = 0
    MAX_RETRIES_PER_PAGE = 2
    while True:
        # Always verify session before reading the result list
        if ensure_logged_in(driver, LOGIN_EMAIL, LOGIN_PASSWORD):
            logger.info("re-login happened — re-opening results")
            if not _open_country_results(driver, country, LOGIN_EMAIL, LOGIN_PASSWORD):
                break
            main_handle = driver.current_window_handle
            for _ in range(page - 1):
                if not click_next_page(driver):
                    break

        page_urls = collect_company_urls(driver)
        if not page_urls:
            logger.info("country %s page %s: no urls — stop", country, page)
            break
        # Filter out urls we've already scraped (in case of a retry)
        page_urls = [u for u in page_urls if u not in seen_urls]
        if not page_urls:
            logger.info("country %s page %s: all urls already done", country, page)
        else:
            if max_companies:
                remaining = max_companies - companies_done
                if remaining <= 0:
                    break
                page_urls = page_urls[:remaining]
            logger.info("country %s page %s: %s urls", country, page, len(page_urls))

            infos, session_lost = _scrape_page_tabs(
                driver, country, page_urls, main_handle, tab_batch=tab_batch,
            )

            if main_handle not in driver.window_handles:
                logger.error("main results tab closed — aborting country")
                break

            if session_lost:
                _close_extra_tabs(driver, main_handle)
                if retries_this_page >= MAX_RETRIES_PER_PAGE:
                    logger.error("page %s exhausted retries — moving on", page)
                else:
                    retries_this_page += 1
                    logger.warning(
                        "session lost on page %s — re-login + retry (%s/%s)",
                        page, retries_this_page, MAX_RETRIES_PER_PAGE,
                    )
                    ensure_logged_in(driver, LOGIN_EMAIL, LOGIN_PASSWORD,
                                     navigate_first=True)
                    if not _open_country_results(driver, country,
                                                  LOGIN_EMAIL, LOGIN_PASSWORD):
                        break
                    main_handle = driver.current_window_handle
                    for _ in range(page - 1):
                        if not click_next_page(driver):
                            break
                    continue  # retry same page

            # Successful batch: write results + remember scraped urls
            retries_this_page = 0
            for info in infos:
                current_index = append_company_rows(ws, current_index, info)
                seen_urls.add(info.url)
                companies_done += 1
                if save_callback and companies_done % save_every == 0:
                    save_callback()

        if max_companies and companies_done >= max_companies:
            logger.info("max_companies hit (%s) — stop", max_companies)
            break
        if max_pages and page >= max_pages:
            logger.info("max_pages hit (%s) — stop", max_pages)
            break
        _close_extra_tabs(driver, main_handle)
        if not click_next_page(driver):
            break
        page += 1

    elapsed = time.time() - started
    if companies_done:
        logger.info(
            "country %s: %s scraped in %.1fs (%.1fs/company)",
            country, companies_done, elapsed, elapsed / companies_done,
        )
    return current_index


def _run_country_serial(driver: WebDriver, country: str, ws, current_index: int,
                        save_every: int = 1, save_callback=None,
                        max_pages: int | None = None,
                        max_companies: int | None = None) -> int:
    """Legacy single-tab flow: collect all URLs then visit one-by-one.
    Kept as a fallback via --no-tabs in case the tabbed flow hits issues.
    """
    urls = collect_all_urls_for_country(
        driver, country, max_pages=max_pages, max_urls=max_companies,
    )
    logger.info("country %s: %s company urls to visit", country, len(urls))
    if not urls:
        return current_index
    if max_companies:
        urls = urls[:max_companies]
        logger.info("max_companies cap: %s", len(urls))

    companies_done = 0
    started = time.time()
    for i, url in enumerate(urls, 1):
        logger.info("  [%s/%s] %s", i, len(urls), url)
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

    elapsed = time.time() - started
    if companies_done:
        logger.info(
            "country %s: %s/%s scraped in %.1fs (%.1fs/company)",
            country, companies_done, len(urls), elapsed, elapsed / companies_done,
        )
    return current_index


def main() -> int:
    parser = argparse.ArgumentParser(description="JCtrans directory scraper")
    parser.add_argument("--country", nargs="*", help="Country name(s) to scrape. Default: all")
    parser.add_argument("--resume", action="store_true", help="Skip countries already in checkpoint")
    parser.add_argument("--max-pages", type=int, default=None, help="Max pages per country (testing)")
    parser.add_argument("--max-companies", type=int, default=None, help="Max companies per country (testing)")
    parser.add_argument("--no-headless", action="store_true", help="Run with visible browser (local only)")
    parser.add_argument("--excel", type=Path, default=EXCEL_PATH, help="Output Excel path")
    parser.add_argument("--no-tabs", action="store_true",
                        help="Disable tab-batching mode (visit detail pages one-by-one)")
    parser.add_argument("--tab-batch", type=int, default=20,
                        help="How many detail pages to open as tabs at once (default 20)")
    parser.add_argument("--browser", choices=["chrome", "edge"], default="chrome",
                        help="Which browser to drive (default: chrome)")
    parser.add_argument(
        "--attach", default=None,
        help=(
            "Attach to an already-running browser via remote-debugging-protocol, "
            "format 'host:port' (e.g. '127.0.0.1:9526'). Use this to drive the "
            "Edge / Chrome window you opened yourself — bot won't launch its own."
        ),
    )
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    setup_logging(args.verbose)
    logger.info("starting jctrans scraper")

    excel_path: Path = args.excel
    wb, ws, next_row = open_or_create_workbook(excel_path)
    state = load_checkpoint()

    driver = make_driver(
        browser=args.browser,
        headless=not args.no_headless,
        attach=args.attach,
    )
    attached = bool(args.attach)
    if attached:
        logger.info("attached to existing %s at %s", args.browser, args.attach)

    def save() -> None:
        wb.save(excel_path)
        save_checkpoint(state)

    try:
        # When attached to user's browser, do not force navigation /
        # credential fill — just check, and only run login() if SIGN IN is
        # actually visible (so we don't disturb a working session).
        if attached:
            ensure_logged_in(driver, LOGIN_EMAIL, LOGIN_PASSWORD, navigate_first=True)
        else:
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
                    use_tabs=not args.no_tabs,
                    tab_batch=args.tab_batch,
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
        if not attached:
            try:
                driver.quit()
            except Exception:
                pass
        else:
            # Detach without killing the user's browser
            try:
                # Best-effort: close every tab we opened, leaving the
                # browser session alive for the user.
                main_handle = driver.current_window_handle
                for h in list(driver.window_handles):
                    if h == main_handle:
                        continue
                    try:
                        driver.switch_to.window(h)
                        driver.close()
                    except Exception:
                        pass
                driver.switch_to.window(main_handle)
            except Exception:
                pass

    logger.info("done — wrote %s", excel_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
