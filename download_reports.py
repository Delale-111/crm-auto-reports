from playwright.sync_api import sync_playwright
import os
import json

LOGIN = os.environ["CRM_LOGIN"]
PASSWORD = os.environ["CRM_PASSWORD"]
DOWNLOAD_DIR = os.path.abspath("./downloads")
HISTORY_FILE = os.path.join(DOWNLOAD_DIR, "downloaded_files.json")
CRM_URL = "https://crm.secureholiday.net/crm/"
REPORTS_URL = "https://crm.secureholiday.net/crm/Dashboards/BiReportExtract/Index/FR"

def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r") as f:
            return set(json.load(f))
    return set()

def save_history(history):
    with open(HISTORY_FILE, "w") as f:
        json.dump(list(history), f)

def main():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    history = load_history()
    new_files = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        page.goto(CRM_URL, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(3000)
        page.fill('input[placeholder="Entrez votre login"]', LOGIN)
        page.fill('input[placeholder="Mot de passe"]', PASSWORD)
        page.click('text=SE CONNECTER')
        page.wait_for_load_state("networkidle", timeout=30000)
        print(f"Connecte : {page.url}")

        page.goto(REPORTS_URL, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(5000)

        download_buttons = page.query_selector_all(
            "a[href*='Download'], a[href*='download'], button.download, "
            ".fa-download, a.btn-download, td a[title], a.glyphicon"
        )
        if not download_buttons:
            download_buttons = page.query_selector_all(
                "table tbody tr td:last-child a, table tbody tr td:last-child button"
            )

        print(f"Trouve {len(download_buttons)} fichiers")

        for i, btn in enumerate(download_buttons):
            try:
                with page.expect_download(timeout=30000) as download_info:
                    btn.click()
                download = download_info.value
                filename = download.suggested_filename

                if filename not in history:
                    filepath = os.path.join(DOWNLOAD_DIR, filename)
                    download.save_as(filepath)
                    new_files.append(filepath)
                    history.add(filename)
                    print(f"Nouveau : {filename}")
                else:
                    print(f"Deja telecharge : {filename}")
            except Exception as e:
                print(f"Erreur {i}: {e}")

        browser.close()

    save_history(history)
    print(f"\nTermine : {len(new_files)} nouveau(x) fichier(s)")

if __name__ == "__main__":
    main()
