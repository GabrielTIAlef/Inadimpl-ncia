from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import WebDriverException, TimeoutException
import os, glob, time, requests, sys

SLACK_WEBHOOK_URL = "https://hooks.slack.com....
DATASET_NAME = "Inadimplência"

def send_slack(msg):
    if not SLACK_WEBHOOK_URL: 
        print(msg); return
    try: requests.post(SLACK_WEBHOOK_URL, json={"text": msg}, timeout=10)
    except Exception as e: print("Slack erro:", e, file=sys.stderr)


USER_DATA_DIR = r"C:\RPA\Chrome\UserData"
PROFILE_DIR = "Default"


CHROME_BINARY = os.getenv("CHROME_BINARY", "").strip()

os.makedirs(USER_DATA_DIR, exist_ok=True)
for lock in glob.glob(os.path.join(USER_DATA_DIR, "Singleton*")):
    try: os.remove(lock)
    except: pass

opts = webdriver.ChromeOptions()
opts.add_argument(f"--user-data-dir={USER_DATA_DIR}")
opts.add_argument(f"--profile-directory={PROFILE_DIR}")
opts.add_argument("--no-first-run")
opts.add_argument("--no-default-browser-check")
opts.add_argument("--disable-blink-features=AutomationControlled")
opts.add_argument("--disable-dev-shm-usage")
opts.add_argument("--remote-allow-origins=*")
opts.add_argument("--start-maximized")
opts.add_argument("--headless=new")
opts.add_argument("--disable-gpu")
opts.add_argument("--window-size=1920,1080")

if CHROME_BINARY:
    opts.binary_location = CHROME_BINARY

service = Service()
driver = None
try:
    driver = webdriver.Chrome(service=service, options=opts)
    driver.get("https://app.powerbi.com/groups/me/list?experience=power-bi")
    wait = WebDriverWait(driver, 40)

    wait.until(EC.presence_of_element_located((By.XPATH, "//a[@data-testid='item-name']")))
    blocks = driver.find_elements(By.XPATH, "//span[contains(@class, 'col-name')]")

    found = False
    for block in blocks:
        try:
            a = block.find_element(By.XPATH, ".//a[@data-testid='item-name']")
            name = (a.get_attribute("aria-label") or "").strip()
            if name == DATASET_NAME:
                ActionChains(driver).move_to_element(block).perform()
                btn = WebDriverWait(block, 10).until(
                    EC.element_to_be_clickable((By.XPATH, ".//button[@data-testid='quick-action-button-Atualizar agora']"))
                )
                btn.click()
                print("'Atualizar agora' clicado!")
                found = True
                break
        except Exception:
            continue

    if not found:
        send_slack(f"Dataset '{DATASET_NAME}'")
    time.sleep(5)

except TimeoutException:
    send_slack("Timeout ao carregar itens do Power BI")
except Exception as e:
    send_slack(f"Erro: {e}")
    raise
finally:
    try: 
        if driver: driver.quit()
    except: 
        pass

