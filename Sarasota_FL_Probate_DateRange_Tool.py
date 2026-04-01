from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd
import os
import re
from datetime import datetime, timedelta

# =========================
# Helpers
# =========================
def norm(s: str) -> str:
    """Normalize text for matching."""
    if s is None:
        return ""
    s = s.strip().upper()
    s = re.sub(r"\s+", " ", s)
    return s

# Priority order (keywords)
PRIORITY_ORDER = [
    "PETITION TO ADMIT WILL",
    "ORDER ADMITTING WILL TO PROBATE",
    "ORDER ADMITTING WILL TO PROBATE AND APPOINTING PERSONAL REPRESENTATIVE",
    "PETITION FOR SUMMARY ADMINISTRATION",
    "PETITION ADMINISTRATION",
    "PETITION FOR ADMINISTRATION",
    "PETITION FOR ADMINISTRATION - ANCILLARY",
    "PETITION FOR ADMINISTRATION FEDERAL",
    "PETITION FOR ANCILLARY ADMINISTRATION",
    "PETITION TO DETERMINE HOMESTEAD STATUS OF REAL PROPERTY",
    "PETITION-ADMINISTRATION",
    "NOTICE OF ADMINISTRATION",
]

DESC_COLS = [f"Description {i}" for i in range(1, 21)]
TESTATE_COL = "Testate Status"
MATCHED_KEYWORD_COL = "Matched Keyword"
MATCHED_DESC_COL = "Matched Description"

def collect_docket_descriptions(driver):
    """Scrape docket grid descriptions from td[4] in original order (unique)."""
    collected = []
    seen = set()

    i = 0
    while True:
        try:
            row_xpath = f'//*[@id="ctl00_cphBody_rgDocket_ctl00__{i}"]'
            row = driver.find_element(By.XPATH, row_xpath)
            desc = row.find_element(By.XPATH, "./td[4]").text.strip()

            if desc:
                dn = norm(desc)
                if dn not in seen:
                    collected.append(desc)
                    seen.add(dn)

            i += 1
        except NoSuchElementException:
            break

    return collected

def compute_testate_status(desc_list):
    """
    Intestate if docket contains:
      - 'WITHOUT A WILL'
      - 'WITHOUT WILL'
    Else Testate if contains 'WILL'
    Else Not Found
    """
    combined = " ".join(norm(d) for d in desc_list if d)

    # Intestate has priority
    if "WITHOUT A WILL" in combined or "WITHOUT WILL" in combined:
        return "Intestate"

    # Then Testate
    if "WILL" in combined:
        return "Testate"

    return "Not Found"

def find_best_priority_match(desc_list):
    """
    Returns (matched_keyword, matched_description) based on contains match.
    - Picks highest PRIORITY_ORDER keyword that appears in ANY description (contains).
    - matched_description is the first description (in docket order) that contains that keyword.
    If no matches -> ("", "")
    """
    normalized = [(d, norm(d)) for d in desc_list if d]

    for kw in PRIORITY_ORDER:
        nkw = norm(kw)
        for original, nd in normalized:
            if nkw in nd:
                return kw, original
    return "", ""

def order_descriptions_priority_first(desc_list, max_out=20):
    """
    Reorder descriptions:
    - First: those matching PRIORITY_ORDER (contains), in PRIORITY_ORDER sequence.
      For each keyword, if multiple docket lines match it, keep them in docket order.
    - Then: remaining descriptions (still in docket order).
    Returns up to max_out.
    """
    normalized = [(d, norm(d)) for d in desc_list if d]

    used = set()  # normalized descriptions
    ordered = []

    # Add priority-matching descriptions first
    for kw in PRIORITY_ORDER:
        nkw = norm(kw)
        for original, nd in normalized:
            if nkw in nd:
                nd0 = norm(original)
                if nd0 not in used:
                    ordered.append(original)
                    used.add(nd0)
                if len(ordered) >= max_out:
                    return ordered[:max_out]

    # Add remaining
    for original, nd in normalized:
        nd0 = norm(original)
        if nd0 not in used:
            ordered.append(original)
            used.add(nd0)
        if len(ordered) >= max_out:
            break

    return ordered[:max_out]

# =========================
# Setup
# =========================
chrome_options = Options()
# ==============================================================
# GitHub Actions Auto-Headless Detection
# ==============================================================
if os.environ.get("GITHUB_ACTIONS") == "true":
    chrome_options.add_argument("--headless=new")
    print("Running in Headless mode (GitHub Actions detected)")
else:
    print("Running in standard mode (Local detected)")

chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

# If you prefer to manually provide the path to chromedriver, you can change the service line.
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

wait = WebDriverWait(driver, 20)
# =========================
# Navigation & Search Setup
# =========================
driver.get("https://secure.sarasotaclerk.com/CaseInfo.aspx")

# 1. Click General Public User Access: Click Here
wait.until(EC.element_to_be_clickable((By.ID, "cphBody_hlAnon"))).click()

# 2. Click Agree
wait.until(EC.element_to_be_clickable((By.ID, "cphBody_bAgree"))).click()

# 3. Click search (in menu)
wait.until(EC.element_to_be_clickable((By.XPATH, '//span[text()="Search"]'))).click()

# 4. Case File Date: (10 days old from today to today)
now = datetime.now()
start_date_obj = now - timedelta(days=10)
start_date = start_date_obj.strftime("%m/%d/%Y").replace("/0", "/").lstrip("0")
end_date = now.strftime("%m/%d/%Y").replace("/0", "/").lstrip("0")

print(f"Searching from {start_date} to {end_date}")

start_input = wait.until(EC.presence_of_element_located((By.ID, "ctl00_cphBody_rdStart_dateInput")))
start_input.clear()
start_input.send_keys(start_date)

end_input = wait.until(EC.presence_of_element_located((By.ID, "ctl00_cphBody_rdEnd_dateInput")))
end_input.clear()
end_input.send_keys(end_date)

# 5. Court Type: Probate/Guardianship/Mental Health
# Open Dropdown using suggested XPath
court_type_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[contains(@id,'rcbCourtType_Input')]")))
court_type_input.click()
time.sleep(1.5)

# Select the Checkbox for Probate/Guardianship/Mental Health
probate_checkbox = wait.until(EC.presence_of_element_located((By.XPATH, "//li[contains(., 'Probate/Guardianship/Mental Health')]//input[@type='checkbox']")))
driver.execute_script("arguments[0].scrollIntoView(true);", probate_checkbox)
driver.execute_script("arguments[0].click();", probate_checkbox)
time.sleep(1)

# 6. Click Probate in Case Type (Multi-Step Logic)
case_type_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[contains(@id,'rcbCaseType_Input')]")))

print("Triggering Case Type refresh (clicking once)...")
case_type_input.click() # First click to trigger refresh

print("Waiting 15 seconds for refresh to complete...")
time.sleep(15) # Wait for page reload as requested

print("Opening Case Type dropdown again...")
# Relocate input in case of DOM refresh
case_type_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[contains(@id,'rcbCaseType_Input')]")))
case_type_input.click() # Second click to open it properly
time.sleep(2)

# Select the Checkbox for Probate in Case Type
probate_case_checkbox = wait.until(EC.presence_of_element_located((By.XPATH, "//li[contains(., 'Probate')]//input[@type='checkbox']")))
driver.execute_script("arguments[0].scrollIntoView(true);", probate_case_checkbox)
driver.execute_script("arguments[0].click();", probate_case_checkbox)
time.sleep(1)

# Close Case Type dropdown
case_type_input.click()
time.sleep(1)

# 7. Click Search
wait.until(EC.element_to_be_clickable((By.ID, "ctl00_cphBody_bSearch_input"))).click()
time.sleep(5)

# Set page size to 50
try:
    dropdown_arrow = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="ctl00_cphBody_rgCaseList_ctl00_ctl03_ctl01_PageSizeComboBox_Arrow"]')
        )
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", dropdown_arrow)
    time.sleep(1)
    dropdown_arrow.click()
    wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="ctl00_cphBody_rgCaseList_ctl00_ctl03_ctl01_PageSizeComboBox_DropDown"]/div/ul/li[4]')
        )
    ).click()
    time.sleep(5)
except Exception as e:
    print(f"Note: Could not set page size to 50 (might already be or only 1 page): {e}")

# =========================
# Output / Resume
# =========================
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
file_path = f"Sarasota_FL_Probate_Output_{timestamp}.xlsx"
existing_cases = set()

if os.path.exists(file_path):
    df_existing = pd.read_excel(file_path)

    # Ensure new columns exist in old file
    for col in [TESTATE_COL, MATCHED_KEYWORD_COL, MATCHED_DESC_COL] + DESC_COLS:
        if col not in df_existing.columns:
            df_existing[col] = ""

    needed_party_cols = [
        'Decedent 1 Name','Decedent 1 Type','Decedent 1 Attorney',
        'Decedent 2 Name','Decedent 2 Type','Decedent 2 Attorney',
        'Applicant 1 Name','Applicant 1 Type','Applicant 1 Attorney',
        'Applicant 2 Name','Applicant 2 Type','Applicant 2 Attorney',
        'Petitioner 1 Name','Petitioner 1 Type','Petitioner 1 Attorney',
        'Petitioner 2 Name','Petitioner 2 Type','Petitioner 2 Attorney',
    ]
    for col in needed_party_cols:
        if col not in df_existing.columns:
            df_existing[col] = ""

    existing_cases = set(df_existing['Case Number'].astype(str))
    data = df_existing.to_dict(orient='records')
    print(f"Loaded {len(existing_cases)} existing records.")
else:
    data = []

# =========================
# Total records & pages
# =========================
total_records = 0
try:
    record_label_elem = wait.until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="cphBody_tCounts"]/tbody/tr/td[1]/label'))
    )
    record_label = record_label_elem.text
    match = re.search(r'(\d+)', record_label)
    if match:
        total_records = int(match.group(1))
except Exception as e:
    print("Could not find record label. 0 records likely.")
    total_records = 0

print(f"Total Records Found: {total_records}")

if total_records == 0:
    print("No records to process. Task Completed.")
    driver.quit()
    exit()

records_per_page = 50
total_pages = (total_records // records_per_page) + (1 if total_records % records_per_page else 0)

def go_to_page(page_num):
    try:
        xpath = f'//*[@id="ctl00_cphBody_rgCaseList_ctl00"]/tfoot/tr/td/table/tbody/tr/td/div[2]/a[{page_num}]'
        wait.until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
        time.sleep(3)
    except Exception as e:
        try:
            next_xpath = '//*[@id="ctl00_cphBody_rgCaseList_ctl00"]/tfoot/tr/td/table/tbody/tr/td/div[3]/input[1]'
            driver.find_element(By.XPATH, next_xpath).click()
            time.sleep(3)
        except:
            print(f"⚠️ Could not navigate to page {page_num}: {e}")

record_index = len(existing_cases) + 1

# =========================
# Column order (FORCE)
# Testate Status after Petitioner 2 Attorney,
# then Matched Keyword, Matched Description,
# then Description 1..20
# =========================
BASE_COLS = [
    "S.No.", "County", "State", "Case Number", "File Date",
    "Decedent 1 Name", "Decedent 1 Type", "Decedent 1 Attorney",
    "Decedent 2 Name", "Decedent 2 Type", "Decedent 2 Attorney",
    "Applicant 1 Name", "Applicant 1 Type", "Applicant 1 Attorney",
    "Applicant 2 Name", "Applicant 2 Type", "Applicant 2 Attorney",
    "Petitioner 1 Name", "Petitioner 1 Type", "Petitioner 1 Attorney",
    "Petitioner 2 Name", "Petitioner 2 Type", "Petitioner 2 Attorney",
    TESTATE_COL, MATCHED_KEYWORD_COL, MATCHED_DESC_COL,
]
ALL_COLS = BASE_COLS + DESC_COLS

def save_excel(data_rows):
    df_out = pd.DataFrame(data_rows)
    for c in ALL_COLS:
        if c not in df_out.columns:
            df_out[c] = ""
    df_out = df_out[ALL_COLS]
    df_out.to_excel(file_path, index=False)

# =========================
# Main scrape loop
# =========================
for page_num in range(1, total_pages + 1):
    if page_num > 1:
        go_to_page(page_num)

    row_index = 0
    while True:
        row_xpath = f'//*[@id="ctl00_cphBody_rgCaseList_ctl00__{row_index}"]'
        try:
            row_elem = driver.find_element(By.XPATH, row_xpath)
            first_td = row_elem.find_element(By.XPATH, './td[1]/a')
            case_text = first_td.text.strip()

            if case_text in existing_cases:
                print(f"[Skipped] Case {case_text} already scraped.")
                row_index += 1
                continue

            driver.execute_script("arguments[0].click();", first_td)
            time.sleep(3)

            case_number = driver.find_element(By.XPATH, '//*[@id="cphBody_CaseNumber"]').text.strip()
            file_date = driver.find_element(By.XPATH, '//*[@id="cphBody_FileDate"]').text.strip()

            # Party extraction
            party_data = {
                'Decedent 1 Name': '', 'Decedent 1 Type': '', 'Decedent 1 Attorney': '',
                'Decedent 2 Name': '', 'Decedent 2 Type': '', 'Decedent 2 Attorney': '',
                'Applicant 1 Name': '', 'Applicant 1 Type': '', 'Applicant 1 Attorney': '',
                'Applicant 2 Name': '', 'Applicant 2 Type': '', 'Applicant 2 Attorney': '',
                'Petitioner 1 Name': '', 'Petitioner 1 Type': '', 'Petitioner 1 Attorney': '',
                'Petitioner 2 Name': '', 'Petitioner 2 Type': '', 'Petitioner 2 Attorney': '',
            }

            party_counts = {'Decedent': 0, 'Applicant': 0, 'Petitioner': 0}

            i = 0
            while True:
                try:
                    type_xpath = f'//*[@id="ctl00_cphBody_rgParty_ctl00__{i}"]/td[2]'
                    row_type = driver.find_element(By.XPATH, type_xpath).text.strip()

                    if row_type in party_counts:
                        prow_xpath = f'//*[@id="ctl00_cphBody_rgParty_ctl00__{i}"]'
                        full_row = driver.find_element(By.XPATH, prow_xpath)
                        cells = full_row.find_elements(By.TAG_NAME, 'td')
                        name = cells[0].text.strip()
                        ptype = cells[1].text.strip()
                        attorney = cells[2].text.strip()

                        count = party_counts[row_type]
                        if count < 2:
                            party_data[f'{row_type} {count + 1} Name'] = name
                            party_data[f'{row_type} {count + 1} Type'] = ptype
                            party_data[f'{row_type} {count + 1} Attorney'] = attorney
                            party_counts[row_type] += 1
                    i += 1
                except NoSuchElementException:
                    break

            # Docket
            docket_descs_raw = collect_docket_descriptions(driver)

            # Testate Status (based on all docket descriptions)
            testate_status = compute_testate_status(docket_descs_raw)

            # Matched Keyword / Matched Description (best priority contains-match)
            matched_keyword, matched_description = find_best_priority_match(docket_descs_raw)

            # Re-ordered Description 1..20
            desc_list = order_descriptions_priority_first(docket_descs_raw, max_out=20)
            desc_data = {f"Description {idx+1}": (desc_list[idx] if idx < len(desc_list) else "") for idx in range(20)}

            case_data = {
                'S.No.': record_index,
                'County': 'Sarasota',
                'State': 'FL',
                'Case Number': case_number,
                'File Date': file_date,
                **party_data,
                TESTATE_COL: testate_status,
                MATCHED_KEYWORD_COL: matched_keyword,
                MATCHED_DESC_COL: matched_description,
                **desc_data,
            }

            data.append(case_data)

            # Save after each case (with forced column order)
            save_excel(data)

            print(f"[{record_index}] Case {case_number} saved. "
                  f"Testate Status: {testate_status} | Matched Keyword: {matched_keyword}")
            existing_cases.add(case_number)
            record_index += 1

            driver.back()
            time.sleep(3)

        except NoSuchElementException:
            break

        row_index += 1

print("Scraping Task Completed. Data saved to Excel.")
driver.quit()
