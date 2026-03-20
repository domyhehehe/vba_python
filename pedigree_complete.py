# coding: utf-8
import importlib.util
import pathlib
import sys
import time

import requests
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


ROOT_DIR = pathlib.Path(__file__).resolve().parent
CSV3_PATH = ROOT_DIR / "csv" / "csv3.py"


def load_csv3_module():
    spec = importlib.util.spec_from_file_location("csv3_complete_base", CSV3_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"failed to load module from {CSV3_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def build_driver(headless: bool = True):
    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--window-size=1600,2200")
    options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
    )
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


def _wait_page_ready(driver, timeout: int):
    wait = WebDriverWait(driver, timeout)
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

    locators = [
        (By.CSS_SELECTOR, "table.pedigreetable"),
        (By.CSS_SELECTOR, "td.w2 a"),
        (By.CSS_SELECTOR, "table[border='1']"),
        (By.TAG_NAME, "body"),
    ]
    for by, value in locators:
        try:
            wait.until(EC.presence_of_element_located((by, value)))
            return
        except TimeoutException:
            continue


def fetch_html_via_browser(url: str, timeout: int = 30, headless: bool = True):
    driver = build_driver(headless=headless)
    try:
        driver.get(url)
        _wait_page_ready(driver, timeout=timeout)
        time.sleep(0.5)
        return driver.page_source
    finally:
        driver.quit()


def install_fetch_overrides(base):
    original_make_session = base.make_session

    def make_session():
        session = original_make_session()
        session.headers.update(
            {
                "Upgrade-Insecure-Requests": "1",
                "Sec-Fetch-Dest": "document",
                "Sec-Fetch-Mode": "navigate",
                "Sec-Fetch-Site": "none",
                "Sec-Fetch-User": "?1",
            }
        )
        return session

    def fetch_html(url: str, session=None, timeout: int = None, prefer_browser: bool = False):
        timeout = timeout or base.DEFAULT_TIMEOUT
        session = session or make_session()
        normalized_url = base._normalize_url(url)

        if not prefer_browser:
            try:
                resp = session.get(normalized_url, timeout=timeout)
                resp.raise_for_status()
                resp.encoding = resp.encoding or resp.apparent_encoding or "utf-8"
                text = resp.text
                if text and ("pedigreetable" in text or 'class="w2"' in text or "query_type=stakes" in text):
                    return text
            except Exception as req_err:
                print(f"[WARN] requests fetch failed: {normalized_url} ({req_err})")

        print(f"[INFO] browser fetch: {normalized_url}")
        return fetch_html_via_browser(normalized_url, timeout=timeout, headless=False)


    def fetch_soup(url: str, session=None, timeout: int = None, prefer_browser: bool = False):
        html = fetch_html(url, session=session, timeout=timeout, prefer_browser=prefer_browser)
        return BeautifulSoup(html, "html.parser")

    base.make_session = make_session
    base.fetch_html = fetch_html
    base.fetch_soup = fetch_soup


def main(argv=None):
    base = load_csv3_module()
    install_fetch_overrides(base)
    argv = list(sys.argv[1:] if argv is None else argv)
    return base.cli_main(argv)


if __name__ == "__main__":
    sys.exit(main())
