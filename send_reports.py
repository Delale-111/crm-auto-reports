import smtplib
import zipfile
import glob
import os
import time
import re
import subprocess
import sys

DOWNLOAD_DIR = os.path.abspath('./downloads')
EMAIL_FROM = os.environ['SMTP_EMAIL']
EMAIL_PASSWORD = os.environ['SMTP_PASSWORD']
EMAIL_TO = os.environ['EMAIL_TO']
SF_ORG = os.environ.get('SF_ORG', 'PROD')

BATCH_SIZE = 3
DELAY_SECONDS = 15


def smtp_connect():
    s = smtplib.SMTP('smtp.office365.com', 587)
    s.starttls()
    s.login(EMAIL_FROM, EMAIL_PASSWORD)
    return s


def find_latest_zip():
    zips = glob.glob(os.path.join(DOWNLOAD_DIR, 'Sunelia_Rapports_indiv_pour_groupe_*.zip'))
    if not zips:
        return None

    def date_key(path):
        m = re.findall(r'\d{4}_\d{2}_\d{2}', os.path.basename(path))
        return m[-1] if m else ''

    return max(zips, key=date_key)


def extract_zip(latest_zip: str) -> str:
    extract_dir = latest_zip.replace('.zip', '')
    if not os.path.exists(extract_dir):
        with zipfile.ZipFile(latest_zip, 'r') as z:
            z.extractall(extract_dir)
        print(f'Dezippe dans : {extract_dir}')
    else:
        print(f'Deja dezippe : {extract_dir}')
    return extract_dir


def generate_eml(excel_path: str, output_dir: str) -> str:
    """Appelle generate_report.py et retourne le chemin du .eml genere."""
    cmd = [
        sys.executable, 'generate_report.py', excel_path,
        '--org', SF_ORG,
        '--output', output_dir,
        '--from', EMAIL_FROM,
        '--to', EMAIL_TO,
    ]
    print(f'  Generation rapport: {os.path.basename(excel_path)}')
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)

    if result.returncode != 0:
        print(f'  STDERR: {result.stderr[-500:]}')
        raise RuntimeError(f'generate_report.py echoue (code {result.returncode})')

    # Trouver le .eml genere
    base = os.path.splitext(os.path.basename(excel_path))[0]
    base_clean = base.replace('Sunelia_Rapports_indiv_pour_groupe_', '')
    eml_path = os.path.join(output_dir, f'Reporting_{base_clean}.eml')

    if not os.path.exists(eml_path):
        # fallback: chercher tout .eml recent
        emls = glob.glob(os.path.join(output_dir, '*.eml'))
        if emls:
            eml_path = max(emls, key=os.path.getmtime)
        else:
            raise FileNotFoundError(f'Aucun .eml trouve dans {output_dir}')

    print(f'  EML genere: {os.path.basename(eml_path)}')
    return eml_path


def send_eml(smtp_conn, eml_path: str, excel_path: str):
    """Envoie un fichier .eml avec le xlsx en piece jointe."""
    from email import policy as epolicy
    from email.mime.base import MIMEBase
    from email import encoders

    with open(eml_path, 'rb') as f:
        msg = email.message_from_bytes(f.read(), policy=epolicy.SMTP)

    with open(excel_path, 'rb') as f:
        part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(excel_path))
        msg.attach(part)

    smtp_conn.sendmail(EMAIL_FROM, [EMAIL_TO], msg.as_bytes())


def main():
    latest_zip = find_latest_zip()
    if not latest_zip:
        print('Aucun zip trouve')
        raise SystemExit('Aucun zip trouve')

    print(f'Zip le plus recent : {os.path.basename(latest_zip)}')
    extract_dir = extract_zip(latest_zip)

    eml_dir = os.path.join(DOWNLOAD_DIR, 'eml_output')
    os.makedirs(eml_dir, exist_ok=True)

    excels_all = sorted(glob.glob(os.path.join(extract_dir, '*.xlsx')))
    SKIP = ['Wecamp', 'Baia']
    excels = [e for e in excels_all if not any(s.lower() in os.path.basename(e).lower() for s in SKIP)]
    print(f'Trouve {len(excels_all)} fichiers Excel, {len(excels)} apres filtrage (exclus: {len(excels_all)-len(excels)})')

    if not excels:
        raise SystemExit('0 Excel trouve dans le zip')

    # Phase 1 : generation de tous les .eml
    eml_files = []
    for excel_path in excels:
        try:
            eml_path = generate_eml(excel_path, eml_dir)
            eml_files.append(eml_path)
        except Exception as e:
            print(f'  ERREUR generation {os.path.basename(excel_path)}: {e}')

    print(f'\n{len(eml_files)} rapports generes sur {len(excels)} Excel')

    if not eml_files:
        raise SystemExit('0 rapport genere')

    # Phase 2 : envoi par lots
    total_sent = 0
    total_errors = 0

    for i in range(0, len(eml_files), BATCH_SIZE):
        batch = eml_files[i:i + BATCH_SIZE]
        print(f'--- Lot {i // BATCH_SIZE + 1} ---')

        s = smtp_connect()
        for eml_path in batch:
            try:
                send_eml(s, eml_path)
                total_sent += 1
                print(f'  Envoye : {os.path.basename(eml_path)} -> {EMAIL_TO}')
            except Exception as e:
                total_errors += 1
                print(f'  ERREUR envoi {os.path.basename(eml_path)}: {e}')
        s.quit()

        if i + BATCH_SIZE < len(eml_files):
            print(f'  Pause {DELAY_SECONDS}s...')
            time.sleep(DELAY_SECONDS)

    print(f'\nTermine ! {total_sent} mails envoyes')
    if total_sent == 0:
        raise SystemExit('0 mails envoyes')
    if total_errors > 0:
        print(f'ATTENTION: {total_errors} erreurs.')


if __name__ == '__main__':
    main()
