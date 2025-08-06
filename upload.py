import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

from automation import (
    login, handle_disclaimer, click_create_new_constituent,
    fill_form, click_next_step, fill_details,
    select_intake_method, click_create_casework, click_create_casework_from_home,
    click_home_button
)


def upload_to_council_connect(df, username, password, auto_click, driver_path):
    """Automates submission of parsed email entries to Council Connect."""

    def element_exists(driver, xpath):
        try:
            driver.find_element(By.XPATH, xpath)
            return True
        except NoSuchElementException:
            return False

    def wait_for_home_screen(driver, wait, timeout=120):
        try:
            WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, "//h2[contains(text(),'Create Casework')]"))
            )
            return True
        except Exception:
            return False

    def set_opened_at_now(driver):
        now = datetime.now()
        formatted_datetime = now.strftime("%Y-%m-%dT%H:%M")
        opened_at_input = driver.find_element(By.ID, "opened_at")
        driver.execute_script("""
            const input = arguments[0];
            const value = arguments[1];
            const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
            nativeInputValueSetter.call(input, value);
            input.dispatchEvent(new Event('input', { bubbles: true }));
        """, opened_at_input, formatted_datetime)

        print(f"🕒 Set 'Opened At' to: {formatted_datetime}")

    LOGIN_URL = "https://councilconnect.council.nyc.gov/login"
    FORM_URL = "https://councilconnect.council.nyc.gov/casework/create"

    FIELD_MAP_STEP_1 = {
        "Name": "newConstituent.name",
        "Email": "newConstituent.contact_info.0.contact_data"
    }

    service = Service(driver_path)
    driver = webdriver.Edge(service=service)
    wait = WebDriverWait(driver, 30)

    try:
        login(driver, wait, LOGIN_URL, username, password)
        handle_disclaimer(driver, wait)

        for i, row in df.iterrows():
            print(f"\n🚀 Submitting entry {i + 1}/{len(df)}")
            driver.get(FORM_URL)
            wait.until(lambda d: d.find_element(By.TAG_NAME, "form"))

            if not click_create_new_constituent(driver, wait):
                print("⏭ Skipping entry: form not opened.")
                continue

            fill_form(driver, row, FIELD_MAP_STEP_1)
            click_next_step(driver, wait)

            wait.until(lambda d: d.find_element(By.ID, "details"))
            if not fill_details(driver, wait, row.get("Reply", "")):
                print("⚠️ Skipped due to error filling details.")
                continue

            click_next_step(driver, wait)
            select_intake_method(driver, wait, "Emailed")
            set_opened_at_now(driver)
            time.sleep(1)

            click_next_step(driver, wait)
            click_next_step(driver, wait)

            if auto_click:
                try:
                    print("⏳ Waiting for 'Create Casework' button to appear...")
                    wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Create Casework')]")))
                    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Create Casework')]")))
                    click_create_casework(driver, wait)
                    time.sleep(4)  # short delay to let form complete
                    print("✅ Casework created. Moving to next person...\n")
                except Exception as e:
                    print(f"❌ Auto-click failed: {e}")
                continue


            else:
                print("🛑 Please click 'Create Casework' manually in the browser...")

                home_screen_loaded = False
                print("⌛ Waiting for user to either submit OR skip (manually return to Home)...")

                while True:
                    if not element_exists(driver, "//button[contains(text(), 'Create Casework')]") and \
                       not element_exists(driver, "//button[contains(text(), 'Next Step')]"):
                        print("✅ Form submitted — detected button disappearance.")
                        break

                    if element_exists(driver, "//h2[contains(text(),'Create Casework')]"):
                        print("⏩ User skipped form — detected return to Home screen.")
                        home_screen_loaded = True
                        time.sleep(2)
                        break

                    time.sleep(4)

        print("\n✅ All entries processed.")

    finally:
        print("🛑 Leaving browser open for inspection.")
        
        driver.quit()