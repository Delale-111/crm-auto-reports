import smtplib
import zipfile
import glob
import os
import time
import re
import io
from email.message import EmailMessage

# --- Optional deps (fallback si pas installées) ---
try:
    import pandas as pd
except Exception:
    pd = None

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

try:
    import matplotlib.pyplot as plt
except Exception:
    plt = None


# --- CONFIG ---
DOWNLOAD_DIR = os.path.abspath("./downloads")
EMAIL_FROM = os.environ["SMTP_EMAIL"]
EMAIL_PASSWORD = os.environ["SMTP_PASSWORD"]
EMAIL_TO_TEST = os.environ["EMAIL_TO"]

# Secret GitHub déjà créé: OPENAI_API_KEY
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")

# Limites (coût / temps)
TAIL_ROWS_PER_SHEET = int(os.environ.get("AI_TAIL_ROWS", "30"))
MAX_SHEETS_ANALYZED = int(os.environ.get("AI_MAX_SHEETS", "12"))

BATCH_SIZE = 3
DELAY_SECONDS = 10

client = None
if OPENAI_API_KEY and OpenAI is not None:
    client = OpenAI(api_key=OPENAI_API_KEY)


def smtp_connect():
    s = smtplib.SMTP("smtp.office365.com", 587)
    s.starttls()
    s.login(EMAIL_FROM, EMAIL_PASSWORD)
    return s


def _strip_code_fences(text: str) -> str:
    return (text or "").replace("```html", "").replace("```", "").strip()


def generer_sparkline_cid(df, cid_prefix="spark"):
    """
    Génère une mini-courbe PNG et retourne (cid, png_bytes).
    IMPORTANT: on utilise CID (<img src="cid:...">) car Outlook bloque souvent data:image;base64.
    """
    if plt is None or df is None:
        return None
    try:
        df_num = df.select_dtypes(include=["number"])
        if df_num.empty or len(df_num) < 2:
            return None

        data = df_num.iloc[-15:, -1].dropna()
        if data.empty:
            return None

        plt.figure(figsize=(2, 0.5))
        plt.plot(data.values, linewidth=2)
        plt.axis("off")
        plt.tight_layout(pad=0)

        buf = io.BytesIO()
        plt.savefig(buf, format="png", transparent=True)
        plt.close()
        buf.seek(0)

        png_bytes = buf.read()
        # CID unique (stable pour un contenu donné)
        cid = f"{cid_prefix}_{abs(hash(png_bytes))}"
        return cid, png_bytes
    except Exception:
        return None


def analyser_feuille_kpi(nom_feuille: str, df) -> str:
    """
    Demande à lIA: KPI + phrase danalyse courte sous forme de 2 <td>.
    """
    if client is None or pd is None or df is None or df.empty:
        return (
            '<td><div style="font-size:18px; font-weight:bold;">N/A</div>'
            '<div style="font-size:10px; color:gray;">Données</div></td>'
            "<td>Pas de données ou IA indisponible.</td>"
        )

    data_str = df.tail(TAIL_ROWS_PER_SHEET).to_csv(index=False)

    prompt = f"""
Tu es un analyste. Voici les données de la feuille '{nom_feuille}'.

Mission : Extraire UN KPI clé et UNE phrase d'analyse pour un email.

Données (CSV) :
{data_str}

Règles strictes :
1) Réponds UNIQUEMENT avec du HTML correspondant à 2 cellules (<td>).
2) Format attendu :
   <td><div style="font-size:18px; font-weight:bold; color:#2c3e50;">[VALEUR KPI]</div>
       <div style="font-size:10px; color:gray;">[Nom KPI]</div></td>
   <td>[Analyse ultra-courte, max 20 mots. Mentionne tendance si possible].</td>
3) Si données vides/illisibles : KPI="N/A" et texte="Pas de données".
4) Ne mets PAS de <tr>.
""".strip()

    try:
        resp = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
        )
        return _strip_code_fences(resp.choices[0].message.content)
    except Exception as e:
        return f"<td>Erreur IA</td><td>{str(e)}</td>"


def generer_intro_globale(raw_kpi_text: str) -> str:
    """
    Intro courte "Bonjour," (3 lignes max).
    """
    if client is None:
        return "Bonjour,<br>Veuillez trouver ci-dessous la synthèse IA du rapport en pièce jointe."

    prompt = f"""
Voici des extraits (KPI + analyses) :
{raw_kpi_text}

Rédige une introduction email (3 lignes max).
Ton : professionnel, direct, positif. Commence par "Bonjour,".
""".strip()

    try:
        resp = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.5,
        )
        return (resp.choices[0].message.content or "").strip().replace("\n", "<br>")
    except Exception:
        return "Bonjour,<br>Veuillez trouver ci-dessous la synthèse IA du rapport en pièce jointe."


def build_ai_dashboard_html(excel_path: str, camping_name: str):
    """
    Retourne (plain_text, html, related_images)
    related_images = list of tuples: (cid, bytes, maintype, subtype)
    """
    if pd is None:
        plain = (
            "Bonjour,\n\nVeuillez trouver ci-joint le rapport.\n"
            "Analyse IA indisponible (pandas non installé).\n\nCordialement,\nSunelia"
        )
        html = f"""
        <html><body style="font-family:Arial,sans-serif">
        <p>Bonjour,<br><br>
        Veuillez trouver ci-joint le rapport pour : <b>{camping_name}</b>.<br>
        <span style="color:#888">Analyse IA indisponible (pandas non installé).</span><br><br>
        Cordialement,<br>Sunelia</p>
        </body></html>
        """
        return plain, html, []

    all_sheets = pd.read_excel(excel_path, sheet_name=None)
    noms = list(all_sheets.keys())

    feuilles = noms[1:-1] if len(noms) >= 3 else noms
    feuilles = feuilles[:MAX_SHEETS_ANALYZED]

    rows_html = ""
    raw_text_for_summary = ""
    related_images = []

    for nom in feuilles:
        df = all_sheets.get(nom)
        if df is None or df.empty:
            continue

        td_ia = analyser_feuille_kpi(nom, df)

        spark = generer_sparkline_cid(df)
        if spark:
            cid, png_bytes = spark
            related_images.append((cid, png_bytes, "image", "png"))
            td_graph = f'<td style="text-align:center;"><img src="cid:{cid}" alt="Sparkline" /></td>'
        else:
            td_graph = '<td style="color:#ccc; font-size:10px;">Pas de graph</td>'

        rows_html += f"""
        <tr style="border-bottom: 1px solid #eee;">
            <td style="padding: 10px; font-weight:bold; color:#d35400;">{nom}</td>
            {td_ia}
            {td_graph}
        </tr>
        """.strip()

        raw_text_for_summary += f"Feuille {nom}: {td_ia}\n"

    intro_html = generer_intro_globale(raw_text_for_summary)

    html = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333; background-color: #f9f9f9; padding: 20px;">
        <div style="max-width: 860px; margin: auto; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.08);">
            <h2 style="color: #2c3e50; border-bottom: 2px solid #e67e22; padding-bottom: 10px;">
                 Synthèse IA  {camping_name}
            </h2>

            <p style="font-size: 14px; line-height: 1.5; color: #555;">
                {intro_html}
            </p>

            <table style="width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 13px;">
                <thead>
                    <tr style="background-color: #f2f2f2; text-align: left;">
                        <th style="padding: 10px; width: 18%;">Sujet</th>
                        <th style="padding: 10px; width: 22%;">KPI Clé</th>
                        <th style="padding: 10px;">Analyse</th>
                        <th style="padding: 10px; width: 18%;">Courbe (15 pts)</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html or '<tr><td colspan="4" style="padding:10px;color:#999;">Aucune donnée exploitable.</td></tr>'}
                </tbody>
            </table>

            <p style="margin-top: 26px; font-size: 11px; color: #999; text-align: center;">
                Rapport généré par IA  détail complet dans le fichier Excel joint.<br>
                Source : {os.path.basename(excel_path)}
            </p>
        </div>
    </body>
    </html>
    """.strip()

    plain = (
        f"Bonjour,\n\nVeuillez trouver ci-joint le rapport pour : {camping_name}\n"
        "Une synthèse IA est disponible en version HTML.\n\n"
        "Cordialement,\nSunelia"
    )
    return plain, html, related_images


def send_one_email(smtp, email_to: str, subject: str, excel_path: str):
    filename = os.path.basename(excel_path)
    camping = (
        filename.replace("Sunelia_Rapports_indiv_pour_groupe_", "")
        .replace(".xlsx", "")
        .replace("_", " ")
    )

    plain_body, html_body, related_images = build_ai_dashboard_html(excel_path, camping)

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = EMAIL_FROM
    msg["To"] = email_to

    msg.set_content(plain_body)
    html_part = msg.add_alternative(html_body, subtype="html")

    # Ajout des images inline via CID (meilleure compatibilité Outlook)
    for cid, data, maintype, subtype in (related_images or []):
        html_part.add_related(data, maintype=maintype, subtype=subtype, cid=cid)

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

    total_sent = 0
    for i in range(0, len(excels), BATCH_SIZE):
        batch = excels[i : i + BATCH_SIZE]
        print(f"--- Lot {i // BATCH_SIZE + 1} ---")

        s = smtp_connect()
        for excel_path in batch:
            try:
                camping = send_one_email(
                    s,
                    EMAIL_TO_TEST,
                    subject=f"Rapport Sunelia - {os.path.basename(excel_path)}",
                    excel_path=excel_path,
                )
                total_sent += 1
                print(f"  Envoye : {camping} -> {EMAIL_TO_TEST}")
            except Exception as e:
                print(f"  ERREUR envoi {os.path.basename(excel_path)}: {e}")
        s.quit()

        if i + BATCH_SIZE < len(excels):
            print(f"  Pause {DELAY_SECONDS}s...")
            time.sleep(DELAY_SECONDS)

    print(f"Termine ! {total_sent} mails envoyes")


if __name__ == "__main__":
    main()
