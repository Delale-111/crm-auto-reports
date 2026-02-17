import smtplib
import zipfile
import glob
import os
import time
import re
from email.message import EmailMessage

DOWNLOAD_DIR = os.path.abspath("./downloads")
EMAIL_FROM = os.environ["SMTP_EMAIL"]
EMAIL_PASSWORD = os.environ["SMTP_PASSWORD"]
EMAIL_TO_TEST = os.environ["EMAIL_TO"]
BATCH_SIZE = 3
DELAY_SECONDS = 30

def send_one_email(smtp, email_to, subject, body, filepath):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = EMAIL_FROM
    msg["To"] = email_to
    msg.set_content(body)
    with open(filepath, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="octet-stream",
            filename=os.path.basename(filepath),
        )
    smtp.send_message(msg)

def main():
    zips = glob.glob(os.path.join(DOWNLOAD_DIR, "Sunelia_Rapports_indiv_pour_groupe_*.zip"))
    if not zips:
        print("Aucun zip trouve")
        return

    def date_key(path):
        m = re.findall(r"\d{4}_\d{2}_\d{2}", os.path.basename(path))
        return m[-1] if m else ""

    latest_zip = max(zips, key=date_key)
    print(f"Zip le plus recent : {os.path.basename(latest_zip)}")

    extract_dir = latest_zip.replace(".zip", "")
    if not os.path.exists(extract_dir):
        with zipfile.ZipFile(latest_zip, "r") as z:
            z.extractall(extract_dir)
        print(f"Dezippe dans : {extract_dir}")
    else:
        print(f"Deja dezippe : {extract_dir}")

    excels = sorted(glob.glob(os.path.join(extract_dir, "*.xlsx")))
    print(f"Trouve {len(excels)} fichiers Excel")

    s = smtplib.SMTP("smtp.office365.com", 587)
    s.starttls()
    s.login(EMAIL_FROM, EMAIL_PASSWORD)

    total_sent = 0
    for i in range(0, len(excels), BATCH_SIZE):
        batch = excels[i : i + BATCH_SIZE]
        print(f"--- Lot {i // BATCH_SIZE + 1} ---")

        for filepath in batch:
            filename = os.path.basename(filepath)
            camping = (
                filename.replace("Sunelia_Rapports_indiv_pour_groupe_", "")
                .replace(".xlsx", "")
                .replace("_", " ")
            )
            body = (
                "Bonjour,\n\nVeuillez trouver ci-joint le rapport pour : "
                + camping
                + "\n\nCordialement,\nSunelia"
            )
            send_one_email(s, EMAIL_TO_TEST, "Rapport Sunelia - " + camping, body, filepath)
            total_sent += 1
            print(f"  Envoye : {camping} -> {EMAIL_TO_TEST}")

        if i + BATCH_SIZE < len(excels):
            print(f"  Pause {DELAY_SECONDS}s avant le prochain lot...")
            time.sleep(DELAY_SECONDS)

    s.quit()
    print(f"Termine ! {total_sent} mails envoyes")

if __name__ == "__main__":
    main()
