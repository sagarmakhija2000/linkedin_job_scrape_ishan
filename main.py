import csv
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional

from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    WebDriverException,
    ElementClickInterceptedException,
)
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

try:
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    ChromeDriverManager = None

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False


PROFILE_DIR = Path("linkedin_chrome_profile")
PROFILE_DIR.mkdir(exist_ok=True)
OUTPUT_FILE_XLSX = Path("linkedin_scrape_results.xlsx")
OUTPUT_FILE_CSV = Path("linkedin_scrape_results.csv")
CONNECTIONS_CSV = Path("connection_profiles_list.csv")
CONNECTIONS_URL = "https://www.linkedin.com/mynetwork/invite-connect/connections/"
HOME_URL = "https://www.linkedin.com"
SCROLL_PAUSE_SECONDS = 2.5
MAX_SCROLL_ATTEMPTS = 10
OUTPUT_COLUMNS = [
    "profile_url",
    "company_url",
    "job_title",
    "job_url",
    "match_text",
    "status",
    "processed",
]


def slow_sleep(min_seconds: float = 2.0, max_seconds: float = 4.0) -> None:
    time.sleep((min_seconds + max_seconds) / 2)


def scroll_profile_until(driver: webdriver.Chrome, check_xpath: str, max_attempts: int = 15, step: int = 600) -> bool:
    """Scroll the profile page gently until an element matching check_xpath appears.
    Returns True if found, False otherwise."""
    for attempt in range(max_attempts):
        try:
            driver.find_element(By.XPATH, check_xpath)
            return True
        except NoSuchElementException:
            # Try scrolling a few candidate containers first, then fallback to window
            scrolled = False
            try:
                # prefer scrolling main if present
                driver.execute_script("var m = document.querySelector('main') || document.scrollingElement; m.scrollBy(0, arguments[0]);", step)
                scrolled = True
            except Exception:
                try:
                    driver.execute_script("window.scrollBy(0, arguments[0]);", step)
                    scrolled = True
                except Exception:
                    scrolled = False

            time.sleep(1.5)
            if not scrolled:
                # as a last resort, try clicking to bring focus and wait
                try:
                    body = driver.find_element(By.TAG_NAME, 'body')
                    body.click()
                except Exception:
                    pass
    return False


def start_browser() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    options.add_argument(f"--user-data-dir={PROFILE_DIR.resolve()}")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    if ChromeDriverManager is not None:
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=options)

    try:
        return webdriver.Chrome(options=options)
    except WebDriverException as exc:
        raise RuntimeError(
            "ChromeDriver could not be started. Install webdriver-manager or add chromedriver to PATH. "
            "Example: pip install webdriver-manager"
        ) from exc


def wait_for_element(driver: webdriver.Chrome, locator, timeout: int = 20):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located(locator))


def wait_for_clickable(driver: webdriver.Chrome, locator, timeout: int = 20):
    return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(locator))


def is_logged_in(driver: webdriver.Chrome) -> bool:
    wait_for_element(driver, (By.TAG_NAME, "body"), timeout=15)
    time.sleep(2)
    if "login" in driver.current_url.lower():
        return False

    login_fields = driver.find_elements(By.ID, "session_key") + driver.find_elements(By.ID, "session_password")
    if login_fields:
        return False

    try:
        driver.find_element(By.XPATH, "//a[contains(@href, '/feed') or contains(@href, '/mynetwork') or contains(@aria-label,'Feed')]")
        return True
    except NoSuchElementException:
        return False


def scroll_to_page_end(driver: webdriver.Chrome) -> None:
    scrollable_selectors = [
        "//div[contains(@class,'mn-connections') or contains(@class,'search-results-container')]",
        "//div[contains(@class,'scaffold-finite-scroll__content')]",
        "//div[contains(@class,'scaffold-finite-scroll__list')]",
        "//main",
        "//body",
    ]

    for sel in scrollable_selectors:
        try:
            element = driver.find_element(By.XPATH, sel)
        except NoSuchElementException:
            continue

        last_scroll_top = -1
        attempts = 0
        for _ in range(MAX_SCROLL_ATTEMPTS * 3):
            try:
                if "body" in sel.lower():
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                else:
                    driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", element)
            except Exception:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

            time.sleep(SCROLL_PAUSE_SECONDS)

            if "body" in sel.lower():
                scroll_top = driver.execute_script("return document.body.scrollHeight")
            else:
                scroll_top = driver.execute_script("return arguments[0].scrollTop;", element)

            if scroll_top == last_scroll_top:
                attempts += 1
                if attempts >= MAX_SCROLL_ATTEMPTS:
                    break
            else:
                last_scroll_top = scroll_top
                attempts = 0

        try:
            if "body" in sel.lower():
                driver.execute_script("window.scrollTo(0, 0);")
            else:
                driver.execute_script("arguments[0].scrollTop = 0;", element)
        except Exception:
            pass
        time.sleep(1.5)
        return

    last_height = driver.execute_script("return document.body.scrollHeight")
    attempts = 0
    while attempts < MAX_SCROLL_ATTEMPTS:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(SCROLL_PAUSE_SECONDS)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            attempts += 1
        else:
            attempts = 0
            last_height = new_height

    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(2)


def collect_connection_profile_urls(driver: webdriver.Chrome) -> List[str]:
    # Collect anchors that look like profile links. We avoid company links.
    anchors = driver.find_elements(By.XPATH, "//a[contains(@href, '/in/') and not(contains(@href, '/company/'))]")
    urls = []
    for anchor in anchors:
        href = anchor.get_attribute("href")
        if href and "/in/" in href and href.startswith("http"):
            href = href.split("?")[0].strip()
            if href not in urls:
                urls.append(href)
    return urls


def save_connection_urls(urls: List[str]) -> None:
    if not urls:
        return
    # write unique ordered list to CONNECTIONS_CSV with a simple visited flag
    existing = []
    if CONNECTIONS_CSV.exists():
        with CONNECTIONS_CSV.open("r", encoding="utf-8", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get("profile_url"):
                    existing.append(row["profile_url"])

    combined = []
    for u in urls:
        if u not in combined and u not in existing:
            combined.append(u)

    # append new ones with visited flag = false
    if combined:
        write_header = not CONNECTIONS_CSV.exists()
        with CONNECTIONS_CSV.open("a", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["profile_url", "saved_at", "visited"])
            if write_header:
                writer.writeheader()
            for u in combined:
                writer.writerow({"profile_url": u, "saved_at": datetime.now(timezone.utc).isoformat(), "visited": "false"})


def load_connection_list() -> List[str]:
    urls: List[str] = []
    if not CONNECTIONS_CSV.exists():
        return urls
    with CONNECTIONS_CSV.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get("profile_url"):
                urls.append(row["profile_url"])
    return urls


def load_processed_profile_urls() -> set:
    processed = set()
    if OUTPUT_FILE_XLSX.exists() and HAS_PANDAS:
        df = pd.read_excel(OUTPUT_FILE_XLSX)
        if "profile_url" in df.columns:
            processed.update(df["profile_url"].dropna().astype(str).tolist())
        elif "job_url" in df.columns:
            processed.update(df["job_url"].dropna().astype(str).tolist())
    elif OUTPUT_FILE_CSV.exists():
        with OUTPUT_FILE_CSV.open("r", encoding="utf-8", newline="") as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                if row.get("profile_url"):
                    processed.add(row["profile_url"])
                elif row.get("job_url"):
                    processed.add(row["job_url"])
    return processed


def append_rows_to_output(rows: List[Dict[str, str]]) -> None:
    if not rows:
        return

    normalized_rows = []
    for row in rows:
        normalized = {key: row.get(key, "") for key in OUTPUT_COLUMNS}
        normalized_rows.append(normalized)

    if OUTPUT_FILE_XLSX.exists() and HAS_PANDAS:
        existing_df = pd.read_excel(OUTPUT_FILE_XLSX)
        combined = pd.concat([existing_df, pd.DataFrame(normalized_rows)], ignore_index=True)
        combined = combined.reindex(columns=OUTPUT_COLUMNS)
        combined.to_excel(OUTPUT_FILE_XLSX, index=False)
        return

    if OUTPUT_FILE_CSV.exists():
        with OUTPUT_FILE_CSV.open("a", encoding="utf-8", newline="") as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=OUTPUT_COLUMNS)
            writer.writerows(normalized_rows)
        return

    with OUTPUT_FILE_CSV.open("w", encoding="utf-8", newline="") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=OUTPUT_COLUMNS)
        writer.writeheader()
        writer.writerows(normalized_rows)


def ensure_output_exists() -> None:
    if HAS_PANDAS and not OUTPUT_FILE_XLSX.exists() and not OUTPUT_FILE_CSV.exists():
        empty = pd.DataFrame(columns=OUTPUT_COLUMNS)
        empty.to_excel(OUTPUT_FILE_XLSX, index=False)
    elif not OUTPUT_FILE_CSV.exists():
        with OUTPUT_FILE_CSV.open("w", encoding="utf-8", newline="") as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=OUTPUT_COLUMNS)
            writer.writeheader()


def scrape_person_profile(driver: webdriver.Chrome, profile_url: str) -> List[Dict[str, str]]:
    info_rows: List[Dict[str, str]] = []
    company_name = ""
    company_url = ""
    status = "unknown"
    profile_window = driver.current_window_handle

    try:
        print(f"Profile tab active: {profile_url}")
        wait_for_element(driver, (By.TAG_NAME, "body"), timeout=15)
        slow_sleep(1.5, 2.5)

        print("Searching for Experience header on profile page...")
        found = scroll_profile_until(driver, "//h2[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'experience')]", max_attempts=12)
        if not found:
            print("Experience section did not load after scrolling.")

        slow_sleep(0.5, 1.5)
        experience_heading = None
        try:
            experience_heading = driver.find_element(
                By.XPATH,
                "//h2[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'experience')]"
            )
        except NoSuchElementException:
            print("Experience heading not found on profile page.")

        company_link = None
        if experience_heading is not None:
            try:
                company_link = experience_heading.find_element(
                    By.XPATH,
                    "./following::a[contains(translate(@href, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'company')][1]"
                )
            except NoSuchElementException:
                company_link = None

        if company_link is None:
            try:
                company_link = driver.find_element(
                    By.XPATH,
                    "(//a[contains(translate(@href, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'company') and contains(@href, '/company/')])[1]"
                )
            except NoSuchElementException:
                company_link = None

        if company_link is None:
            print("No company link under Experience; skipping profile.")
            status = "no current company found in Experience"
            info_rows.append({
                "profile_url": profile_url,
                "company_url": "",
                "job_title": "",
                "job_url": "",
                "match_text": "",
                "status": status,
                "processed": "true",
            })
            return info_rows

        company_name = company_link.text.strip() or "(company name not found)"
        company_url = company_link.get_attribute("href").split("?")[0].strip()
        status = "company found"
        print(f"Found company in Experience: {company_name} -> {company_url}")

        print(f"Opening company page in a new tab: {company_url}")
        driver.execute_script("window.open(arguments[0]);", company_url)
        company_window = driver.window_handles[-1]
        driver.switch_to.window(company_window)
        wait_for_element(driver, (By.TAG_NAME, "body"), timeout=15)
        slow_sleep(1.5, 2.5)
        print("Company page loaded.")

        company_jobs_link = None
        try:
            company_jobs_link = driver.find_element(
                By.XPATH,
                "//a[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'jobs') and contains(@href,'/company/') and contains(@href,'/jobs')]"
            )
            print(f"Company Jobs link found: {company_jobs_link.get_attribute('href')}")
        except NoSuchElementException:
            print("No company-specific Jobs link found on company page.")

        if company_jobs_link is None:
            info_rows.append({
                "profile_url": profile_url,
                "company_url": "",
                "job_title": "",
                "job_url": "",
                "match_text": "",
                "status": "no company-specific jobs tab",
                "processed": "true",
            })
            return info_rows

        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", company_jobs_link)
            slow_sleep(0.5, 1.0)
            print("Clicking company Jobs link...")
            try:
                company_jobs_link.click()
            except ElementClickInterceptedException:
                try:
                    ActionChains(driver).move_to_element(company_jobs_link).pause(0.5).click(company_jobs_link).perform()
                except Exception:
                    driver.execute_script("arguments[0].click();", company_jobs_link)
            time.sleep(1)
        except Exception as exc:
            print(f"Failed to click company Jobs link: {exc}")
            slow_sleep(1, 2)

        try:
            time.sleep(2)
            jobs_slider = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//ul[contains(@class,'artdeco-carousel__slider') and contains(@class,'ember-view')]")
                )
            )
        except TimeoutException:
            print("No artdeco-carousel__slider found on the jobs page.")
            info_rows.append({
                "profile_url": profile_url,
                "company_url": "",
                "job_title": "",
                "job_url": "",
                "match_text": "",
                "status": "company jobs carousel not found",
                "processed": "true",
            })
            return info_rows

        job_items = jobs_slider.find_elements(By.XPATH, ".//li")
        print(f"Found {len(job_items)} carousel items to process.")

        matched_jobs = []
        seen_urls = set()

        for idx, item in enumerate(job_items, start=1):
            try:
                job_anchor = None
                href = ""

                for candidate in item.find_elements(By.XPATH, ".//a[@href]"):
                    candidate_href = candidate.get_attribute("href")
                    if not candidate_href:
                        continue
                    candidate_href = candidate_href.strip()
                    if candidate_href in ("#", "about:blank"):
                        continue

                    lower = candidate_href.lower()
                    is_real_job_url = (
                        "/jobs/view/" in lower
                        or ("/jobs/search-results/" in lower and "currentjobid=" in lower)
                        or ("/jobs/search-results/" in lower and "current_job_id=" in lower)
                        or ("/jobs/search-results/" in lower and "currentjobid" in lower)
                    )

                    if is_real_job_url:
                        job_anchor = candidate
                        href = candidate_href
                        break

                if job_anchor is None:
                    continue

                if href in seen_urls:
                    continue
                seen_urls.add(href)

                try:
                    job_title = job_anchor.find_element(By.XPATH, ".//strong").text.strip()
                except NoSuchElementException:
                    job_title = (job_anchor.text or item.text or "").strip()

                print(f"Opening job {idx}/{len(job_items)} in new tab: {href}")

                driver.execute_script("window.open(arguments[0]);", href)
                handles = driver.window_handles
                driver.switch_to.window(handles[-1])
                try:
                    wait_for_element(driver, (By.TAG_NAME, "body"), timeout=12)
                    slow_sleep(1.0, 1.5)

                    match_text = ""
                    try:
                        time.sleep(2)
                        match_paragraphs = WebDriverWait(driver, 8).until(
                            lambda d: [
                                el for el in d.find_elements(By.TAG_NAME, "p")
                                if "your profile and resume" in el.text.lower()
                            ]
                        )
                        if match_paragraphs:
                            match_text = match_paragraphs[0].text.strip()
                            print(f"Found profile text on job page: {match_text}")
                        else:
                            print("No paragraph matching 'Your profile and resume' found.")
                    except TimeoutException:
                        print("Timed out waiting for the match paragraph to appear.")
                    except Exception:
                        print("Could not read the matching profile paragraph.")

                    filtered_match_text = match_text.strip()
                    filtered_text = filtered_match_text.lower()
                    if filtered_text in {"your profile and resume are missing some required qualifications", "your profile and resume match some required qualifications"}:
                        print(f"Filtering out generic match text for job: {href}")
                        continue

                    matched_jobs.append({
                        "profile_url": profile_url,
                        "company_url": company_url,
                        "job_title": job_title,
                        "job_url": href,
                        "match_text": filtered_match_text,
                        "status": "matched job" if filtered_match_text else "no match text found",
                        "processed": "true",
                    })
                finally:
                    driver.close()
                    driver.switch_to.window(company_window)
                    slow_sleep(0.5, 1.0)
            except Exception as exc:
                print(f"Error inspecting carousel item: {exc}")
                try:
                    driver.switch_to.window(company_window)
                except Exception:
                    pass
                continue

        if not matched_jobs:
            print("No matching job cards found for this company.")
            info_rows.append({
                "profile_url": profile_url,
                "company_url": "",
                "job_title": "",
                "job_url": "",
                "match_text": "",
                "status": "no matching jobs cards found",
                "processed": "true",
            })
            driver.close()
            driver.switch_to.window(profile_window)
            return info_rows

        driver.close()
        driver.switch_to.window(profile_window)
        return matched_jobs

    except Exception as exc:
        print(f"Error while processing profile {profile_url}: {exc}")
        info_rows.append({
            "profile_url": profile_url,
            "company_url": "",
            "job_title": "",
            "job_url": "",
            "match_text": str(exc),
            "status": "error",
            "processed": "false",
        })
        return info_rows


def main() -> None:
    ensure_output_exists()
    processed = load_processed_profile_urls()

    print(f"Starting LinkedIn scraper with profile folder: {PROFILE_DIR.resolve()}")
    print("Using output file:", OUTPUT_FILE_XLSX if HAS_PANDAS else OUTPUT_FILE_CSV)

    driver = start_browser()
    try:
        driver.get(HOME_URL)
        time.sleep(3)
        if not is_logged_in(driver):
            print("LinkedIn login not detected.")
            print("Please login manually in the opened browser, then restart this script.")
            time.sleep(50)
            return

        driver.get(CONNECTIONS_URL)
        wait_for_element(driver, (By.TAG_NAME, "body"), timeout=20)
        slow_sleep(3, 4)
        print("Scrolling through connections to load all profiles...")
        scroll_to_page_end(driver)

        connection_urls = collect_connection_profile_urls(driver)
        save_connection_urls(connection_urls)
        # Load consolidated connection list (including previously saved)
        connection_urls = load_connection_list()
        print(f"Found {len(connection_urls)} connection profile URLs (saved to {CONNECTIONS_CSV}).")

        main_window = driver.current_window_handle

        for index, profile_url in enumerate(connection_urls, 1):
            if profile_url in processed:
                print(f"[{index}/{len(connection_urls)}] Skipping already traced profile: {profile_url}")
                continue

            print(f"[{index}/{len(connection_urls)}] Opening profile in new tab: {profile_url}")
            driver.execute_script("window.open(arguments[0]);", profile_url)
            handles = driver.window_handles
            driver.switch_to.window(handles[-1])

            try:
                rows = scrape_person_profile(driver, profile_url)
                append_rows_to_output(rows)
                processed.add(profile_url)
                print(f"Saved {len(rows)} rows for profile: {profile_url}")
            finally:
                driver.close()
                driver.switch_to.window(main_window)
                slow_sleep(3, 5)

    finally:
        print("Closing browser.")
        driver.quit()


if __name__ == "__main__":
    main()
