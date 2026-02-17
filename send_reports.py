import win32com.client
import zipfile
import glob
import os
import time
import re

# === CONFIG TEST ===
DOWNLOAD_DIR = os.path.abspath("./downloads")
EMAIL_TO_TEST = "delalemerveille@gmail.com"
BATCH_SIZE = 3
DELAY_SECONDS = 30  # pause entre chaque lot

# 1. Trouver le zip le plus recent
zips = glob.glob(os.path.join(DOWNLOAD_DIR, "Sunelia_Rapports_indiv_pour_groupe_*.zip"))
if not zips:
    print("Aucun zip trouve")
    exit()

latest_zip = max(zips, key=lambda f: re.findall(r'\d{4}_\d{2}_\d{2}', f)[-1])
print(f"Zip le plus recent : {os.path.basename(latest_zip)}")

# 2. Dezipper
extract_dir = latest_zip.replace(".zip", "")
if not os.path.exists(extract_dir):
    with zipfile.ZipFile(latest_zip, 'r') as z:
        z.extractall(extract_dir)
    print(f"Dezippe dans : {extract_dir}")
else:
    print(f"Deja dezippe : {extract_dir}")

# 3. Lister les fichiers Excel
excels = sorted(glob.glob(os.path.join(extract_dir, "*.xlsx")))
print(f"Trouve {len(excels)} fichiers Excel\n")

# 4. Envoyer par lots de 3
outlook = win32com.client.Dispatch("Outlook.Application")
total_sent = 0

for i in range(0, len(excels), BATCH_SIZE):
    batch = excels[i:i + BATCH_SIZE]
    print(f"--- Lot {i // BATCH_SIZE + 1} ---")

    for filepath in batch:
        filename = os.path.basename(filepath)
        # Extraire le nom du camping du nom de fichier
        camping = filename.replace("Sunelia_Rapports_indiv_pour_groupe_", "").replace(".xlsx", "").replace("_", " ")

        mail = outlook.CreateItem(0)
        mail.To = EMAIL_TO_TEST
        mail.Subject = f"Rapport Sunelia - {camping}"
        mail.Body = (
            f"Bonjour,\n\n"
            f"Veuillez trouver ci-joint le rapport pour : {camping}\n\n"
            f"Cordialement,\n"
            f"Sunelia"
        )
        mail.Attachments.Add(os.path.abspath(filepath))
        mail.Send()
        total_sent += 1
        print(f"  Envoye : {camping} -> {EMAIL_TO_TEST}")

    if i + BATCH_SIZE < len(excels):
        print(f"  Pause {DELAY_SECONDS}s avant le prochain lot...")
        time.sleep(DELAY_SECONDS)

print(f"\nTermine ! {total_sent} mails envoyes")