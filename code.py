import sys
import time
import os
import csv
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout
from tqdm import tqdm

def get_column_matching(df, substring):
    for col in df.columns:
        if substring.lower() in str(col).lower():
            return col
    return None

def scrape_reps(file_path, output_filepath='scraped_doctors.csv'):
    # Read the center names from Excel/CSV
    data = []
    
    try:
        print(f"Reading data from {file_path}...")
        if str(file_path).endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            # Install xlrd for .xls or openpyxl for .xlsx
            df = pd.read_excel(file_path)
            
        center_col = get_column_matching(df, 'Nombre de Centro')
        province_col = get_column_matching(df, 'Provincia') # Finds "Provincia"
                        
        if not center_col:
            print("Error: Could not find a column containing 'Nombre de Centro' in the file.")
            print("Columns available:", list(df.columns))
            return
            
        for index, row in df.iterrows():
            c = str(row[center_col]).strip() if pd.notna(row[center_col]) else ""
            p = str(row[province_col]).strip() if province_col and pd.notna(row[province_col]) else ""
            
            if c and c != "nan":
                data.append({"center": c, "province": p})
                    
    except Exception as e:
        print(f"Failed to read file {file_path}: {e}")
        return
    
    if not data:
        print("No center names found to search.")
        return

    print(f"Found {len(data)} centers to search.")
    scraped_data = []

    with sync_playwright() as p:
        # Launch Chromium once
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        page.set_default_timeout(60000) # 60s timeout

        for item in tqdm(data, desc="Scraping Centers"):
            center_name = item['center']
            province_name = item['province']
            
            print(f"---")
            print(f"Searching for center: {center_name} in {province_name}")
            
            try:
                page.goto("https://reps.sanidad.gob.es/reps-web/inicio.htm", wait_until="load", timeout=60000)
                time.sleep(3) # Wait for page and dropdowns to fully initialize
                
                # Expand filter if needed
                filter_toggle = page.locator('a[href="#collapseFiltro"]')
                if filter_toggle.count() > 0 and filter_toggle.is_visible():
                    if filter_toggle.get_attribute("aria-expanded") == "false":
                        filter_toggle.click()
                        time.sleep(0.5)

                # 1. Update Province in Dropdown if provided
                if province_name and province_name != "nan":
                    # Case insensitive search through options
                    prov_select = page.locator('select#provincia_filtro')
                    options = prov_select.locator('option')
                    count = options.count()
                    matched_value = None
                    for i in range(count):
                        opt_text = options.nth(i).inner_text().strip().lower()
                        # Normalizing strings to remove accents could be an improvement here
                        if province_name.lower() in opt_text or opt_text in province_name.lower():
                            matched_value = options.nth(i).get_attribute("value")
                            break
                            
                    if matched_value:
                        prov_select.select_option(value=matched_value)
                        print(f"  Selected province matching: {province_name}")
                    else:
                        print(f"  Warning: Province '{province_name}' not found locally in dropdown options.")

                # 2. Enter Center Name
                page.fill('input#centro_filtro', center_name)

                # 3. Click Filter
                page.click('a#filtro')
                time.sleep(4) # Delay for data table below to populate via AJAX

                # Wait for table
                try:
                    page.wait_for_selector('#profesionalTable tbody tr', timeout=30000)
                except:
                    print(f"  No results found for {center_name}.")
                    continue

                page_num = 1
                while True:
                    rows = page.locator('#profesionalTable tbody tr')
                    res_count = rows.count()
                    print(f"  Processing page {page_num} with {res_count} rows.")

                    for i in range(res_count):
                        try:
                            row = page.locator('#profesionalTable tbody tr').nth(i)
                            btn = row.locator('button.btn-primary')
                            if not btn.is_visible():
                                continue
                            
                            btn.click()
                            time.sleep(3) # Wait for modal to fade in
                            
                            info_tab = page.locator('div#informacion')
                            personal_info = info_tab.inner_text() if info_tab.count() > 0 else "N/A"

                            academic_link = page.locator('a[href="#academicos"]')
                            if academic_link.count() > 0 and academic_link.is_visible():
                                academic_link.click()
                                time.sleep(2)
                            acad_tab = page.locator('div#academicos')
                            academic_info = acad_tab.inner_text() if acad_tab.count() > 0 else "N/A"

                            prof_link = page.locator('a[href="#situacionProfesional"]')
                            if prof_link.count() > 0 and prof_link.is_visible():
                                prof_link.click()
                                time.sleep(2)
                            prof_tab = page.locator('div#situacionProfesional')
                        
                            professional_info = "N/A"
                            professional_centers = "N/A"
                        
                            if prof_tab.count() > 0:
                                professional_info = prof_tab.inner_text()
                            
                                centers_data = []
                                cat_rows = page.locator('#situacionProfesionalTable tbody tr')
                                for c in range(cat_rows.count()):
                                    row = cat_rows.nth(c)
                                    try:
                                        tds = row.locator('td')
                                        if tds.count() == 0:
                                            continue
                                        cls = tds.first.get_attribute("class") or ""    
                                        if "dataTables_empty" in cls:
                                            continue
                                        
                                        cat_name = tds.nth(0).inner_text().strip()
                                        row.click()
                                        time.sleep(2)
                                    
                                        centros_link = page.locator('#datosSituacionProfesional a').filter(has_text="Centros")
                                        if centros_link.count() > 0:
                                            if centros_link.first.get_attribute("aria-expanded") != "true":
                                                centros_link.first.click()
                                                time.sleep(2)
                                        
                                            c_rows = page.locator('#centroSituacionTable tbody tr')
                                            centers = []
                                            for idx in range(c_rows.count()):
                                                ctr_row = c_rows.nth(idx)
                                                ctds = ctr_row.locator('td')
                                                if ctds.count() == 2:
                                                    cls = ctds.first.get_attribute("class") or ""
                                                    if "dataTables_empty" not in cls:
                                                        nm = ctds.nth(0).inner_text().strip()
                                                        mun = ctds.nth(1).inner_text().strip()
                                                        centers.append(f"{nm} - {mun}")
                                            if centers:
                                                centers_data.append(f"[{cat_name}] " + ", ".join(centers))
                                    except Exception as e:
                                        print(f"    Error reading category {c}: {e}")
                                    
                                if centers_data:
                                    professional_centers = " | ".join(centers_data)
                                    
                            print(f"  -> Extracted Centers: {professional_centers}")

                            new_row = {
                                "Searched Center Name": center_name,
                                "Searched Province": province_name,
                                "Personal Info": personal_info.strip().replace('\n', ' | '),
                                "Academic Info": academic_info.strip().replace('\n', ' | '),
                                "Professional Info": professional_info.strip().replace('\n', ' | '),
                                "Professional Centers": professional_centers.strip().replace('\n', ' | ')
                            }
                            scraped_data.append(new_row)
                            
                            # Save incrementally so data is not lost
                            try:
                                file_exists = os.path.isfile(output_filepath)
                                with open(output_filepath, 'a', newline='', encoding='utf-8') as output_file:
                                    dict_writer = csv.DictWriter(output_file, fieldnames=new_row.keys())
                                    if not file_exists:
                                        dict_writer.writeheader()
                                    dict_writer.writerow(new_row)
                            except Exception as e:
                                print(f"  [Warning] Failed to append to CSV (is the file open in Excel?): {e}")


                            close_btn = page.locator('button.close, .modal-header .close, button.btn-default, button[data-dismiss="modal"]')
                            if close_btn.count() > 0 and close_btn.first.is_visible():
                                close_btn.first.click()
                                time.sleep(1)
                            else:
                                page.keyboard.press("Escape")
                                time.sleep(1)
                            
                        except Exception as inner_e:
                            print(f"  Error reading row {i}: {inner_e}")
                            # Force reload page if we're broken inside a modal
                            page.reload()
                            time.sleep(1)
                            break
                        
                    # Check for Next page
                    try:
                        next_btn = page.locator('#profesionalTable_next')
                        if next_btn.count() > 0 and next_btn.is_visible():
                            cls = next_btn.get_attribute("class") or ""
                            if "disabled" not in cls:
                                print("  Clicking next page...")
                                next_btn.locator('a').first.click()
                                time.sleep(4)
                                page.wait_for_selector('#profesionalTable tbody tr', timeout=30000)
                                page_num += 1
                                continue
                    except Exception as e:
                        print(f"  Error navigating pagination: {e}")
                    
                    break # Break out of while True loop if no next page

            except PlaywrightTimeout as t:
                print(f"Timeout occurred reading {center_name}: {t}")
            except Exception as e:
                print(f"Error occurred with {center_name}. Event loop might be closed: {e}")
                # Re-create page if context died
                try:
                    page.close()
                except:
                    pass
                try:
                    page = context.new_page()
                    page.set_default_timeout(60000)
                except:
                    # Context or browser is completely dead. We attempt to relaunch browser.
                    browser = p.chromium.launch(headless=False)
                    context = browser.new_context()
                    page = context.new_page()
                    page.set_default_timeout(60000)
                    pass
                
        # Write to csv
        if scraped_data:
            keys = scraped_data[0].keys()
            with open(output_filepath, 'w', newline='', encoding='utf-8') as output_file:
                dict_writer = csv.DictWriter(output_file, fieldnames=keys)
                dict_writer.writeheader()
                dict_writer.writerows(scraped_data)
            print(f"\nSuccess: Saved scraped data to {output_filepath}")
        else:
            print("\nNo doctor details could be scraped.")

        browser.close()

if __name__ == "__main__":
    input_file = r"C:\Roche_Projects\Research_Automation\listadoCentros.xls"
    if len(sys.argv) >= 2:
        input_file = sys.argv[1]
    
    scrape_reps(input_file)