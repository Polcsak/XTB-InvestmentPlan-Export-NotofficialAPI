import os
import time
import shutil
import csv
import re
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
import subprocess
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# Konfigurácia ciest a nastavení
chrome_path = r"C:\Users\jakub\Testing_Chrome\chrome-win64\chrome.exe"
chrome_debug_port = "9222"
user_data_path = r"C:\Users\jakub\Testing_Chrome\chrome_selenium_profile"
csv_file_path = r"C:\Users\jakub\SynologyDrive\osobna_zlozka\Projects\Overviews\Finance_Overview\Data\Balance\accounts.csv"

# 1. Spustenie Chrome v debug móde s použitím špecifického testovacieho profilu
chrome_command = f'"{chrome_path}" --remote-debugging-port={chrome_debug_port} --user-data-dir="{user_data_path}" --disable-component-update'
subprocess.Popen(chrome_command, shell=True)
time.sleep(3)


# 5. Pripojenie na bežiaci Chrome a načítanie stránky BudgetBakers
options = webdriver.ChromeOptions()
options.debugger_address = f"127.0.0.1:{chrome_debug_port}"
driver = webdriver.Chrome(options=options)
driver.get("https://xstation5.xtb.com")
time.sleep(5)



# 3. Klikni na tlačidlo "Login"
login_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "input.xs-btn-ok-login"))
)
login_button.click()

print("✅ Prihlásenie prebehlo (ak boli správne údaje).")

time.sleep(10)

import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#KLIKNUTIE EXPORT 1
try:
    export_button = driver.execute_script('''
        return document
            .querySelector("#contentLayer0 > div.xs-hdivided-container > div:nth-child(2) > div.xs-tab-module-container.xs-tab-module-container-active > div > div > xs6-history")
            .shadowRoot.querySelector("xs6-history-feature > div > div.tab-bar-container > button");
    ''')
    
    if export_button:
        driver.execute_script("arguments[0].click();", export_button)
        print("✅ Klikol som na tlačidlo Export.")
    else:
        print("❌ Tlačidlo Export sa nenašlo.")
except Exception as e:
    print("❌ Chyba pri kliknutí na tlačidlo Export:", e)

#KLIKNUTIE NA DATUMOVE POLE
time.sleep(2)

try:
    date_input = driver.execute_script('''
        return document.querySelector("#contentLayer0 > div.xs-hdivided-container > div:nth-child(2) > div.xs-tab-module-container.xs-tab-module-container-active > div > div > xs6-history")
            .shadowRoot.querySelector("xs6-history-feature > div > div.history-tabs-container > div > xs6-history_featureclosedpositions")
            .shadowRoot.querySelector("xs6-closed-positions-feature > div > xs6-export-report-dialog > dialog > div > section > div.pds-modal__content > div:nth-child(1) > xs6-date-select > div > pds-datepicker-input-from > pds-control > div > input");
    ''')
    
    driver.execute_script("arguments[0].click();", date_input)
    print("✅ Klikol som na dátumové pole.")
except Exception as e:
    print("❌ Chyba pri kliknutí na dátumové pole:", e)

#KLIKNUTIE NA "ALL"
time.sleep(2)

try:
    all_option = driver.execute_script('''
        return document.querySelector("#contentLayer0 > div.xs-hdivided-container > div:nth-child(2) > div.xs-tab-module-container.xs-tab-module-container-active > div > div > xs6-history")
            .shadowRoot.querySelector("xs6-history-feature > div > div.history-tabs-container > div > xs6-history_featureclosedpositions")
            .shadowRoot.querySelector("#cdk-menu-0 > pds-action-list-item:nth-child(7)");
    ''')
    
    if all_option:
        driver.execute_script("arguments[0].click();", all_option)
        print("✅ Klikol som na možnosť 'All' records.")
    else:
        print("❌ Možnosť 'All' records sa nenašla.")
except Exception as e:
    print("❌ Chyba pri kliku na možnosť 'All' records:", e)

#KLIKNUTIE EXPORT 2 (FINAL)
time.sleep(2)
try:
    export_button = driver.execute_script('''
        return document.querySelector("#contentLayer0 > div.xs-hdivided-container > div:nth-child(2) > div.xs-tab-module-container.xs-tab-module-container-active > div > div > xs6-history")
            .shadowRoot.querySelector("xs6-history-feature > div > div.history-tabs-container > div > xs6-history_featureclosedpositions")
            .shadowRoot.querySelector("xs6-closed-positions-feature > div > xs6-export-report-dialog > dialog > div > footer > div > button.pds-button.pds-button--size--l.pds-button--style--primary");
    ''')
    if export_button:
        driver.execute_script("arguments[0].click();", export_button)
        print("✅ Klikol som na tlačidlo Exportovať report.")
    else:
        print("❌ Export button sa nenašiel.")
except Exception as e:
    print("❌ Chyba pri kliku na Export button:", e)


#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
time.sleep(15)
# Ukončenie len procesov spustených s profilom chrome_selenium_profile (tieto diagnostické hlášky budú skryté)
os.system('wmic process where "commandline like \'%%chrome_selenium_profile%%\'" call terminate >nul 2>&1')
print("Automatizovaná inštancia Chrome bola zatvorená.")

# Cesta do priečinka sťahovania (Downloads)
downloads_folder = r"C:\Users\jakub\Downloads"
target_folder = r"C:\Users\jakub\SynologyDrive\osobna_zlozka\Projects\Overviews\Finance_Overview\Data\Investments"

time.sleep(3)

# 4. Vyhľadaj a presuň súbor
for filename in os.listdir(downloads_folder):
    if filename.startswith("account_50266286") and filename.endswith(".xlsx"):
        source_path = os.path.join(downloads_folder, filename)
        target_path = os.path.join(target_folder, "xtb_export.xlsx")
        shutil.move(source_path, target_path)
        print(f"✅ Súbor bol presunutý a premenovaný na: {target_path}")
        break
else:
    print("❌ Súbor sa nenašiel v priečinku Downloads.")
