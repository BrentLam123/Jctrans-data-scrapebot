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
            # Snapshot the first card href so we can wait until the list refreshes
            first_href_before = ""
            try:
                first_link = driver.find_element(
                    By.CSS_SELECTOR,
                    "ul.membership-list-content-center-list > li a[href*='/en/company/']",
                )
                first_href_before = first_link.get_attribute("href") or ""
            except NoSuchElementException:
                pass
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            driver.execute_script("arguments[0].click();", btn)
            try:
                WebDriverWait(driver, 25).until(
                    lambda d: (
                        (lambda links: bool(links) and (links[0].get_attribute("href") or "") != first_href_before)(
                            d.find_elements(
                                By.CSS_SELECTOR,
                                "ul.membership-list-content-center-list > li a[href*='/en/company/']",
                            )
                        )
                    )
                )
            except TimeoutException:
                logger.warning("timeout waiting for next page to refresh")
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


def _parse_current_tab(driver: WebDriver, country: str, url: str,
                       wait_secs: int = 25) -> CompanyInfo | None:
    """Parse the detail page that is already loaded in the current tab.

    Does NOT call driver.get(). Returns None if the company is suspended /
    non-member.
    """
    # Wait for the company-name <p> to appear so we don't sleep blindly
    try:
        WebDriverWait(driver, wait_secs).until(
            lambda d: bool(
                d.find_elements(By.XPATH, "//main//p[contains(@class,'font-bold')]")
            )
            or bool(
                d.find_elements(
                    By.XPATH,
                    "//*[contains(.,'It is NOT a JCtrans member') or contains(.,'MEMBERSHIP SUSPEND')]",
                )
            )
        )
    except TimeoutException:
        logger.warning("detail page slow / never finished: %s", url)

    if is_suspended(driver):
        logger.info("skip suspended/non-member: %s", url)
        return None

    # Jump straight to the bottom to trigger lazy-load (Vue IntersectionObserver).
    # One big jump is much faster than multiple small scrolls and works just as
    # well in practice — the contactCard WebDriverWait below catches the result.
    driver.execute_script(
        "window.scrollTo(0, document.body.scrollHeight);"
    )
    close_blocking_overlays(driver)

    try:
        WebDriverWait(driver, 6).until(
            lambda d: bool(d.find_elements(By.CSS_SELECTOR, "div.contactCard"))
        )
    except TimeoutException:
        # Fallback: try a couple of incremental scrolls in case the single jump
        # didn't trigger the observer (some pages render only after a real
        # scroll step).
        for y in (1500, 4000, 6000):
            driver.execute_script(f"window.scrollTo(0, {y});")
            time.sleep(0.1)
        try:
            WebDriverWait(driver, 4).until(
                lambda d: bool(d.find_elements(By.CSS_SELECTOR, "div.contactCard"))
            )
        except TimeoutException:
            pass
    driver.execute_script("window.scrollTo(0, 0);")

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
    """Open each URL in a new background tab, return the list of new handles in order."""
    before = set(driver.window_handles)
    for url in urls:
        driver.execute_script("window.open(arguments[0], '_blank');", url)
        if stagger:
            time.sleep(stagger)
    # Wait until all tabs are present
    try:
        WebDriverWait(driver, 15).until(
            lambda d: len(set(d.window_handles) - before) >= len(urls)
        )
    except TimeoutException:
        logger.warning(
            "opened only %s/%s tabs", len(set(driver.window_handles) - before), len(urls)
        )
    # Preserve the order in which they were opened
    new_handles: list[str] = []
    seen = set(before)
    for h in driver.window_handles:
        if h not in seen:
            new_handles.append(h)
            seen.add(h)
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
            # Detect session loss BEFORE wasting time parsing masked data
            if not is_logged_in(driver):
                logger.warning("logged out (detected on %s) — aborting batch", url)
                session_lost = True
                try:
                    driver.close()
                except Exception:
                    pass
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
