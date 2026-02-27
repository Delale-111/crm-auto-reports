import smtplib
import zipfile
import glob
import os
import time
import re
import io
from email.message import EmailMessage

# Optional deps
try:
    import pandas as pd
except Exception:
    pd = None

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
except Exception:
    plt = None

try:
    from PIL import Image
except Exception:
    Image = None


DOWNLOAD_DIR = os.path.abspath("./downloads")
EMAIL_FROM = os.environ["SMTP_EMAIL"]
EMAIL_PASSWORD = os.environ["SMTP_PASSWORD"]
EMAIL_TO = os.environ["EMAIL_TO"]

BATCH_SIZE = 3
DELAY_SECONDS = 10

# Animation settings
MAX_POINTS = int(os.environ.get("SPARK_POINTS", "15"))
FRAME_MS = int(os.environ.get("SPARK_FRAME_MS", "140"))  # vitesse animation


def smtp_connect():
    s = smtplib.SMTP("smtp.office365.com", 587)
    s.starttls()
    s.login(EMAIL_FROM, EMAIL_PASSWORD)
    return s


def find_latest_zip():
    zips = glob.glob(os.path.join(DOWNLOAD_DIR, "Sunelia_Rapports_indiv_pour_groupe_*.zip"))
    if not zips:
        return None

    def date_key(path):
        m = re.findall(r"\d{4}_\d{2}_\d{2}", os.path.basename(path))
        return m[-1] if m else ""

    return max(zips, key=date_key)


def extract_zip(latest_zip: str) -> str:
    extract_dir = latest_zip.replace(".zip", "")
    if not os.path.exists(extract_dir):
        with zipfile.ZipFile(latest_zip, "r") as z:
            z.extractall(extract_dir)
        print(f"Dezippe dans : {extract_dir}")
    else:
        print(f"Deja dezippe : {extract_dir}")
    return extract_dir


def pick_timeseries_from_excel(excel_path: str):
    """
    Trouve une série numérique simple dans le classeur :
    - parcourt les feuilles
    - prend la dernière colonne numérique trouvée
    - retourne les N derniers points
    """
    if pd is None:
        return None

    try:
        xls = pd.ExcelFile(excel_path)
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            if df is None or df.empty:
                continue
            num = df.select_dtypes(include=["number"])
            if num is None or num.empty:
                continue
            s = num.iloc[:, -1].dropna()
            if s.empty:
                continue
            values = s.tail(MAX_POINTS).tolist()
            if len(values) >= 3:
                return values
        return None
    except Exception:
        return None


def make_animated_gif(values):
    """
    Génère un GIF animé (courbe qui se trace progressivement).
    Retourne bytes du GIF.
    """
    if plt is None or Image is None or not values or len(values) < 3:
        return None

    frames = []
    n = len(values)

    # normalisation simple (évite un axe complètement plat)
    vmin = min(values)
    vmax = max(values)
    if vmax == vmin:
        vmax = vmin + 1.0

    for k in range(2, n + 1):
        fig = plt.figure(figsize=(2.2, 0.7))
        ax = fig.add_subplot(111)
        ax.plot(values[:k], linewidth=2)
        ax.set_xlim(0, n - 1)
        ax.set_ylim(vmin, vmax)
        ax.axis("off")
        fig.tight_layout(pad=0)

        buf = io.BytesIO()
        fig.savefig(buf, format="png", transparent=True, dpi=110)
        plt.close(fig)
        buf.seek(0)
        frames.append(Image.open(buf).convert("RGBA"))

    out = io.BytesIO()
    frames[0].save(
        out,
        format="GIF",
        save_all=True,
        append_images=frames[1:],
        duration=FRAME_MS,
        loop=0,
        disposal=2,
        optimize=True,
    )
    return out.getvalue()


def build_email(excel_path: str):
    filename = os.path.basename(excel_path)
    camping = (
        filename.replace("Sunelia_Rapports_indiv_pour_groupe_", "")
        .replace(".xlsx", "")
        .replace("_", " ")
    )

    values = pick_timeseries_from_excel(excel_path)
    gif_bytes = make_animated_gif(values) if values else None

    plain = (
        f"Bonjour,\n\nVeuillez trouver ci-joint le rapport pour : {camping}.\n"
        "Une courbe (GIF) est incluse si votre client mail l'anime.\n\n"
        "Cordialement,\nSunelia"
    )

    # NOTE: certains clients (Outlook desktop) n'animent pas les GIF -> 1ère frame
    if gif_bytes:
        cid = f"trend_{abs(hash(gif_bytes))}"
        html = f"""
        <html><body style="font-family:Arial,sans-serif">
          <h2 style="margin:0 0 10px 0;">📈 Synthèse — {camping}</h2>
          <p>Bonjour,<br>Veuillez trouver ci-dessous la courbe (GIF) et le fichier Excel en pièce jointe.</p>
          <div style="margin:10px 0;">
            <img src="cid:{cid}" alt="Courbe animée" />
          </div>
          <p style="font-size:11px;color:#777;">
            Si la courbe semble figée, votre client mail ne supporte pas l'animation GIF.
          </p>
        </body></html>
        """.strip()
        related = [(cid, gif_bytes, "image", "gif")]
    else:
        html = f"""
        <html><body style="font-family:Arial,sans-serif">
          <h2 style="margin:0 0 10px 0;">📈 Synthèse — {camping}</h2>
          <p>Bonjour,<br>Veuillez trouver le fichier Excel en pièce jointe.</p>
          <p style="font-size:11px;color:#777;">Courbe non disponible (pas de données numériques exploitables).</p>
        </body></html>
        """.strip()
        related = []

    subject = f"Rapport Sunelia - {camping}"
    return subject, camping, plain, html, related


def send_one_email(smtp, excel_path: str):
    subject, camping, plain_body, html_body, related_images = build_email(excel_path)

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO

    msg.set_content(plain_body)
    msg.add_alternative(html_body, subtype="html")

    # récupérer la partie HTML et la passer en multipart/related
    html_part = msg.get_payload()[-1]
    try:
        html_part.make_related()
    except Exception:
        pass

    for cid, data, maintype, subtype in (related_images or []):
        html_part.add_related(data, maintype=maintype, subtype=subtype, cid=f"<{cid}>")

    # attache Excel
    with open(excel_path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="octet-stream",
            filename=os.path.basename(excel_path),
        )

    smtp.send_message(msg)
    return camping


def main():
    latest_zip = find_latest_zip()
    if not latest_zip:
        print("Aucun zip trouve")
        raise SystemExit("Aucun zip trouve")

    print(f"Zip le plus recent : {os.path.basename(latest_zip)}")
    extract_dir = extract_zip(latest_zip)

    excels = sorted(glob.glob(os.path.join(extract_dir, "*.xlsx")))
    print(f"Trouve {len(excels)} fichiers Excel")

    if not excels:
        raise SystemExit("0 Excel trouve dans le zip")

    total_sent = 0
    total_errors = 0

    for i in range(0, len(excels), BATCH_SIZE):
        batch = excels[i : i + BATCH_SIZE]
        print(f"--- Lot {i // BATCH_SIZE + 1} ---")

        s = smtp_connect()
        for excel_path in batch:
            try:
                camping = send_one_email(s, excel_path)
                total_sent += 1
                print(f"  Envoye : {camping} -> {EMAIL_TO}")
            except Exception as e:
                total_errors += 1
                print(f"  ERREUR envoi {os.path.basename(excel_path)}: {e}")
        s.quit()

        if i + BATCH_SIZE < len(excels):
            print(f"  Pause {DELAY_SECONDS}s...")
            time.sleep(DELAY_SECONDS)

    print(f"Termine ! {total_sent} mails envoyes")
    if total_sent == 0:
        raise SystemExit("0 mails envoyes (voir erreurs ci-dessus)")
    if total_errors > 0:
        print(f"ATTENTION: {total_errors} erreurs sur l'envoi.")


if __name__ == "__main__":
    main()
