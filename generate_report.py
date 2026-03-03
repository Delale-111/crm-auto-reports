"""
Sunêlia Report Generator (EMAIL .eml) — UI THEME "Sunêlia teal" + benchmark premium (speedometers)
───────────────────────────────────────────────────────────────────────────────────────────────
- Génère un email .eml HTML email-safe
- Graphiques en PNG inline (CID) : montée en charge, ventes semaine, donut produits, speedometers benchmark

Dépendances:
    pip install pandas openpyxl requests matplotlib

Usage:
    python generate_report.py "chemin/vers/rapport.xlsx"
    python generate_report.py "chemin/vers/rapport.xlsx" --org PROD --output "C:/Reports"
    python generate_report.py "chemin/vers/rapport.xlsx" --dry-run

Options email:
    --from "reporting@sunelia.com"
    --to "destinataire@exemple.com"
    --subject "Reporting Sunêlia - ..."
"""

import sys, json, subprocess, argparse, re, time, shutil, math
from pathlib import Path
from io import BytesIO
from html import escape as hesc
from typing import Optional, List, Tuple

try:
    import pandas as pd
except ImportError:
    print("❌ pandas requis. Installe-le : pip install pandas openpyxl")
    sys.exit(1)

try:
    import requests
except ImportError:
    print("❌ requests requis. Installe-le : pip install requests")
    sys.exit(1)

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter
    from matplotlib.patches import Wedge, FancyBboxPatch, Circle
except ImportError:
    print("❌ matplotlib requis pour générer les graphiques email. Installe-le : pip install matplotlib")
    sys.exit(1)

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email import policy


# ═══════════════════════════════════════════════════════
# THEME (Sunêlia-inspired teal / lagon)
# ═══════════════════════════════════════════════════════
THEME = {
    # brand-ish teal palette (derived from Sunêlia visuals)
    "teal_900": "#00465d",   # deep petrol
    "teal_800": "#00596a",
    "teal_700": "#007279",
    "teal_600": "#028f8c",
    "teal_500": "#01b4a3",   # bright lagoon

    "bg":       "#eef7f7",   # global email background
    "surface":  "#ffffff",
    "panel":    "#f3fbfb",   # light aqua panel
    "border":   "#dbe9ea",

    "text":     "#073041",   # deep blue/teal text
    "muted":    "#6f8590",   # muted teal-grey
    "muted2":   "#8ea3ab",

    # status colors (kept readable & premium)
    "good":     "#0aa37a",
    "warn":     "#f2b640",
    "info":     "#1c6e8c",
    "bad":      "#d64545",

    # highlight accents
    "accent":   "#01b4a3",
    "sun":      "#f2b640",   # warm accent (sun)
}


# ═══════════════════════════════════════════════════════
# 1. EXTRACTION EXCEL
# ═══════════════════════════════════════════════════════

SHEETS_CONFIG = {
    "Vue synthétique":        {"max_rows": 60},
    "Suivi activite globale": {"max_rows": 70},
    "Ventes produit":         {"max_rows": 50},
    "Ventes sem.":            {"max_rows": 110},
    "Montée en charge":       {"max_rows": 40},
    "Bassins emetteurs":      {"max_rows": 110},
    "Couverture":             {"max_rows": 10},
}


def find_sheet(available: List[str], target: str) -> Optional[str]:
    target_lower = target.lower()
    for s in available:
        if s.lower() == target_lower:
            return s
    prefix = target[:12].lower()
    for s in available:
        if s.lower().startswith(prefix):
            return s
    return None


def sheet_to_text(df: pd.DataFrame, max_rows: int = 70) -> str:
    lines = []
    for i, row in df.iterrows():
        if i >= max_rows:
            break
        vals = []
        for v in row.values:
            if pd.notna(v):
                s = str(v).strip()
                if s and s != "NaN":
                    if len(s) > 120:
                        s = s[:120] + "..."
                    vals.append(s)
        if vals:
            lines.append(f"R{i}| {' | '.join(vals)}")
    return "\n".join(lines)


def extract_excel(filepath: str) -> str:
    xls = pd.ExcelFile(filepath)
    available = xls.sheet_names
    blocks = []

    for target_name, cfg in SHEETS_CONFIG.items():
        sheet = find_sheet(available, target_name)
        if not sheet:
            continue
        print(f"        → Feuille: {sheet}")
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        text = sheet_to_text(df, cfg["max_rows"])
        blocks.append(f"{'='*40}\nFEUILLE: {sheet}\n{'='*40}\n{text}\n")

    return "\n".join(blocks)


# ═══════════════════════════════════════════════════════
# 2. PROMPT + SCHÉMA JSON
# ═══════════════════════════════════════════════════════

JSON_SCHEMA = r"""
{
  "meta": {
    "camping_name": "string",
    "date_observation": "string ex: 23/02/2026",
    "date_comparaison": "string ex: 24/02/2025"
  },
  "kpi_fermes": {
    "ca":               {"n": number, "n1": number, "var_pct": number},
    "sejours":          {"n": number, "n1": number, "var_pct": number},
    "nuits":            {"n": number, "n1": number, "var_pct": number},
    "prix_moyen_nuit":  {"n": number, "n1": number, "var_pct": number},
    "taux_occupation":  {"n": number, "n1": number, "delta_pts": number},
    "revpar":           {"n": number, "n1": number, "var_pct": number}
  },
  "kpi_total": {
    "ca":      {"n": number, "var_pct": number},
    "sejours": {"n": number, "var_pct": number},
    "stock":   number
  },
  "benchmark": {
    "ca":      {"camping": number, "region": number, "reseau": number},
    "sejours": {"camping": number, "region": number, "reseau": number},
    "region_label": "string ex: Région France",
    "region_nb":    number,
    "reseau_label": "string ex: Réseau Sunêlia",
    "reseau_nb":    number
  },
  "montee_charge": [
    {"date": "string court ex: 31/08", "n": number, "n1": number}
  ],
  "ventes_semaine": [
    {"sem": "string ex: 04/07-10/07", "n": number, "n1": number}
  ],
  "produits": {
    "location":    {"ca": number, "var_pct": number, "sejours_n": number, "pmn": number},
    "emplacement": {"ca": number, "var_pct": number, "sejours_n": number, "pmn": number},
    "total_ca":    number
  },
  "saisonnalite": [
    {"periode": "string", "ca_n": number, "ca_n1": number, "var_pct": number, "sejours": number, "taux_occ": number, "is_haute": boolean}
  ],
  "bassins": [
    {"region": "string", "ca": number, "pct_total": number, "var_pct": number}
  ],
  "insights": [
    {"type": "positive|warning|info|alert", "title": "string 4-6 mots", "text": "string 1-2 phrases max"}
  ]
}
"""


def build_prompt(data_text: str) -> str:
    return f"""Tu es un analyste expert en hôtellerie de plein air pour le réseau Sunêlia.

DONNÉES BRUTES EXTRAITES DU RAPPORT EXCEL:
---
{data_text}
---

MISSION:
Analyse ces données en profondeur. Extrais tous les indicateurs clés et retourne UNIQUEMENT un objet JSON valide (sans texte avant/après, sans backticks markdown) qui suit exactement ce schéma :

{JSON_SCHEMA}

RÈGLES STRICTES:
1. Retourne UNIQUEMENT le JSON. Rien d'autre. Pas de texte, pas de ```, pas d'explication.
2. Les montants en euros arrondis à l'entier (pas de décimales sauf prix_moyen_nuit et taux_occupation).
3. Les pourcentages de variation sont des nombres (ex: 27.6 pour +27,6%). Négatif si baisse.
4. Le delta taux d'occupation en points de pourcentage (ex: -0.57).
5. montee_charge: toutes les dates de la feuille "Montée en charge" — utilise le CA cumulé hors résidentiel.
6. ventes_semaine: toutes les semaines de "Ventes sem." — section hors résidentiel uniquement.
7. saisonnalite: toutes les périodes (Basse saison 1, Moyenne saison, Haute saison 1, Très haute saison, Haute saison 2, Dernière semaine août, Basse saison 2). Les données viennent de "Vue synthétique", section réservations fermes en bas du tableau.
8. bassins: top 8 régions françaises par CA décroissant + si présents les pays étrangers significatifs. Inclure le pct_total = ca_region / ca_total_france * 100.
9. produits: Location = somme des mobil-homes/cottages. Emplacement = emplacements nus. Depuis "Ventes produit".
10. benchmark: les variations camping / région / réseau proviennent de "Suivi activite globale" ou "Vue synthétique" (colonnes VAR région et VAR réseau).
11. insights: EXACTEMENT 6 insights. Types variés (au moins 1 positive, 1 warning ou alert). Basés UNIQUEMENT sur les données. Pas d'invention. Concis. En français.
12. kpi_fermes = ligne "HORS PRODUITS RÉSIDENTIELS" de la section "Réservations fermes" de "Vue synthétique".
13. kpi_total = ligne "HORS PRODUITS RÉSIDENTIELS" de la section "Options et réservations".
14. Le nom du camping se trouve dans la feuille Couverture ou dans "Suivi activite globale".

JSON:"""


# ═══════════════════════════════════════════════════════
# 3. APPEL API SALESFORCE EINSTEIN
# ═══════════════════════════════════════════════════════

def get_sf_auth(org: str = "PROD") -> tuple[str, str]:
    sf_path = shutil.which("sf")
    if not sf_path:
        raise FileNotFoundError("Commande 'sf' introuvable depuis Python.")

    print(f"        → {sf_path} org display --target-org {org}")
    result = subprocess.run(
        [sf_path, "org", "display", "--target-org", org, "--json"],
        capture_output=True, text=True, encoding="utf-8"
    )
    if result.returncode != 0:
        raise RuntimeError(f"Erreur sf cli: {result.stderr}")

    data = json.loads(result.stdout)
    return data["result"]["accessToken"], data["result"]["instanceUrl"]


def call_einstein(prompt: str, token: str, instance_url: str) -> str:
    url = f"{instance_url}/services/apexrest/einstein/generate"
    payload = {"prompt": prompt, "model": "sfdc_ai__DefaultBedrockAnthropicClaude4Sonnet"}
    resp = requests.post(
        url,
        json=payload,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json; charset=utf-8"},
        timeout=180
    )
    resp.raise_for_status()
    return resp.text


def clean_json(raw: str) -> str:
    raw = raw.strip()
    data = raw

    for _ in range(5):
        if isinstance(data, str):
            data = data.strip()
            data = re.sub(r'^\s*```json?\s*', '', data)
            data = re.sub(r'\s*```\s*$', '', data)
            data = data.strip()
            try:
                data = json.loads(data)
            except json.JSONDecodeError:
                break

        if isinstance(data, dict) and "text" in data:
            data = data["text"]
            continue

        if isinstance(data, dict) and "meta" in data:
            return json.dumps(data, ensure_ascii=False)

    if isinstance(data, dict) and "meta" in data:
        return json.dumps(data, ensure_ascii=False)

    raise ValueError(
        f"Impossible d'extraire le JSON d'analyse. Type final: {type(data).__name__}, Aperçu: {str(data)[:200]}"
    )


# ═══════════════════════════════════════════════════════
# 4. FORMAT + HTML EMAIL-SAFE
# ═══════════════════════════════════════════════════════

def fmt_fr(v, decimals=0) -> str:
    try:
        n = float(v)
    except Exception:
        return "-"
    s = f"{n:,.{decimals}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", " ")


def fmt_eur(v) -> str:
    return f"{fmt_fr(v,0)} €"


def fmt_pct(v, decimals=1, signed=True) -> str:
    try:
        n = float(v)
    except Exception:
        return "-"
    sign = "+" if signed and n > 0 else ""
    return f"{sign}{fmt_fr(n,decimals)}%"


def fmt_pts(v, decimals=2) -> str:
    try:
        n = float(v)
    except Exception:
        return "-"
    sign = "+" if n > 0 else ""
    return f"{sign}{fmt_fr(n,decimals)} pts"


def badge_html(v, mode="pct") -> str:
    try:
        n = float(v)
    except Exception:
        n = 0.0

    if n > 0:
        bg, fg, arrow = "#e8fbf5", THEME["good"], "▲"
    elif n < 0:
        bg, fg, arrow = "#ffecec", THEME["bad"], "▼"
    else:
        bg, fg, arrow = "#edf3f4", THEME["muted"], "="

    txt = fmt_pct(n) if mode == "pct" else fmt_pts(n)
    return (
        f"<span style=\"display:inline-block;font-size:11px;font-weight:800;"
        f"padding:3px 9px;border-radius:999px;background:{bg};color:{fg};\">"
        f"{arrow} {hesc(txt)}</span>"
    )


def yk(meta_date: str) -> str:
    return meta_date[-4:] if isinstance(meta_date, str) and len(meta_date) >= 4 else ""


def kpi_cell(label, value, sub, badge) -> str:
    return f"""
    <td style="vertical-align:top;border:1px solid {THEME['border']};padding:12px 12px;border-radius:12px;background:{THEME['surface']};">
      <div style="font-size:10px;text-transform:uppercase;letter-spacing:1px;color:{THEME['muted2']};font-weight:800;margin-bottom:6px;">{hesc(label)}</div>
      <div style="font-size:20px;font-weight:900;color:{THEME['text']};line-height:1.1;margin-bottom:4px;">{hesc(value)}</div>
      <div style="font-size:11px;color:{THEME['muted']};margin-bottom:8px;">{hesc(sub)}</div>
      {badge}
    </td>
    """.strip()


def section_title(title: str) -> str:
    return f"""
    <tr><td style="padding:18px 22px 8px;">
      <div style="font-size:13px;font-weight:900;color:{THEME['text']};text-transform:uppercase;letter-spacing:1.3px;
                  border-bottom:2px solid {THEME['accent']};padding-bottom:8px;">
        {hesc(title)}
      </div>
    </td></tr>
    """.strip()


def small_note(text: str) -> str:
    return f"""
    <tr><td style="padding:0 22px 10px;color:{THEME['muted']};font-size:12px;">
      {hesc(text)}
    </td></tr>
    """.strip()


# ═══════════════════════════════════════════════════════
# 5. GRAPHIQUES PNG (email-safe)
# ═══════════════════════════════════════════════════════

def _mpl_setup():
    plt.rcParams["font.family"] = "DejaVu Sans"
    plt.rcParams["axes.titleweight"] = "bold"
    plt.rcParams["axes.edgecolor"] = THEME["border"]
    plt.rcParams["axes.labelcolor"] = THEME["text"]
    plt.rcParams["xtick.color"] = THEME["muted"]
    plt.rcParams["ytick.color"] = THEME["muted"]


def plot_montee_charge(data: dict) -> bytes:
    _mpl_setup()
    x = [d.get("date", "") for d in data.get("montee_charge", [])]
    y = [d.get("n", 0) for d in data.get("montee_charge", [])]
    y1 = [d.get("n1", 0) for d in data.get("montee_charge", [])]

    fig = plt.figure(figsize=(9.2, 3.2), dpi=170)
    fig.patch.set_facecolor("white")
    ax = fig.add_subplot(111)

    ax.plot(x, y, linewidth=2.6, marker="o", markersize=3.6,
            label="Saison N", color=THEME["teal_500"])
    ax.fill_between(range(len(y)), y, [0]*len(y), color=THEME["teal_500"], alpha=0.08)

    ax.plot(x, y1, linewidth=2.0, linestyle="--", marker="o", markersize=3.0,
            label="Saison N-1", color=THEME["teal_900"])

    ax.grid(True, axis="y", alpha=0.22)
    ax.set_title("Montée en charge — CA cumulé", fontsize=11, color=THEME["text"])
    ax.tick_params(axis="x", labelrotation=45, labelsize=8)
    ax.tick_params(axis="y", labelsize=8)

    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{int(v/1000)}k€"))
    ax.legend(loc="upper left", fontsize=8, frameon=False)

    fig.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def plot_ventes_semaine(data: dict) -> bytes:
    _mpl_setup()
    x = [d.get("sem", "") for d in data.get("ventes_semaine", [])]
    y = [d.get("n", 0) for d in data.get("ventes_semaine", [])]
    y1 = [d.get("n1", 0) for d in data.get("ventes_semaine", [])]

    fig = plt.figure(figsize=(9.2, 3.4), dpi=170)
    fig.patch.set_facecolor("white")
    ax = fig.add_subplot(111)

    idx = list(range(len(x)))
    w = 0.38
    ax.bar([i - w/2 for i in idx], y, width=w, label="Saison N", color=THEME["teal_500"])
    ax.bar([i + w/2 for i in idx], y1, width=w, label="Saison N-1", color="#b4d1d4")

    ax.set_xticks(idx)
    ax.set_xticklabels(x, rotation=45, ha="right", fontsize=8)
    ax.tick_params(axis="y", labelsize=8)
    ax.grid(True, axis="y", alpha=0.22)
    ax.set_title("CA par semaine d'occupation", fontsize=11, color=THEME["text"])
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{int(v/1000)}k€"))
    ax.legend(loc="upper left", fontsize=8, frameon=False)

    fig.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def plot_produits_donut(data: dict) -> bytes:
    _mpl_setup()
    p = data.get("produits", {})
    loc = float(p.get("location", {}).get("ca", 0) or 0)
    emp = float(p.get("emplacement", {}).get("ca", 0) or 0)

    fig = plt.figure(figsize=(4.4, 4.0), dpi=170)
    fig.patch.set_facecolor("white")
    ax = fig.add_subplot(111)
    ax.set_title("Répartition CA par produit", fontsize=11, color=THEME["text"])

    vals = [max(loc, 0.0), max(emp, 0.0)]
    labels = ["Location", "Emplacements"]
    colors = [THEME["teal_500"], THEME["sun"]]

    wedges, _ = ax.pie(vals, startangle=90, colors=colors, wedgeprops=dict(width=0.38))
    total = sum(vals) if sum(vals) > 0 else 1.0

    ax.text(0, 0.05, f"{int((loc + emp) / 1000)} k€", ha="center", va="center",
            fontsize=12, fontweight="bold", color=THEME["text"])
    ax.text(0, -0.12, "CA total", ha="center", va="center", fontsize=9, color=THEME["muted"])

    ax.legend(
        wedges,
        [f"{labels[i]} ({vals[i]/total*100:.1f}%)" for i in range(2)],
        loc="lower center",
        bbox_to_anchor=(0.5, -0.08),
        ncol=1,
        frameon=False,
        fontsize=9
    )

    fig.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


# ──────────────────────────────────────────────────────
# PREMIUM BENCHMARK : Speedometers semi-circulaires
# ──────────────────────────────────────────────────────

def nice_scale_max(values: List[float]) -> float:
    mx = max([abs(float(v)) for v in values] + [0.0])
    if mx <= 8:
        return 10.0
    if mx <= 16:
        return 20.0
    if mx <= 24:
        return 30.0
    if mx <= 35:
        return 40.0
    if mx <= 45:
        return 50.0
    return float(int((mx + 9) // 10) * 10)


def _draw_speedometer(ax, value: float, label: str, ring_color: str, scale_max: float):
    try:
        v = float(value)
    except Exception:
        v = 0.0

    v = max(-scale_max, min(scale_max, v))
    angle = 90.0 - (v / scale_max) * 90.0  # -max=180, 0=90, +max=0

    ax.set_aspect("equal")
    ax.axis("off")
    ax.set_xlim(-1.2, 1.2)
    ax.set_ylim(-0.25, 1.25)

    # Card background (premium)
    card = FancyBboxPatch(
        (-1.15, -0.22), 2.30, 1.45,
        boxstyle="round,pad=0.02,rounding_size=0.14",
        linewidth=0.9, edgecolor=THEME["border"], facecolor=THEME["surface"]
    )
    ax.add_patch(card)

    # Inner panel
    panel = FancyBboxPatch(
        (-1.05, -0.12), 2.10, 1.25,
        boxstyle="round,pad=0.02,rounding_size=0.12",
        linewidth=0.0, facecolor=THEME["panel"]
    )
    ax.add_patch(panel)

    r = 0.95
    width = 0.18

    # Background arc
    bg = Wedge((0, 0), r, 0, 180, width=width, facecolor="#dfecee", edgecolor="none")
    ax.add_patch(bg)

    # Value arc from 0 baseline (90°)
    if v >= 0:
        theta1, theta2 = angle, 90.0
    else:
        theta1, theta2 = 90.0, angle

    arc = Wedge((0, 0), r, theta1, theta2, width=width, facecolor=ring_color, edgecolor="none")
    ax.add_patch(arc)

    # Ticks
    ax.text(-r, -0.02, f"-{int(scale_max)}%", ha="left", va="top", fontsize=7.5, color=THEME["muted2"])
    ax.text(0,  r+0.02, "0", ha="center", va="bottom", fontsize=8.5, color=THEME["muted2"], fontweight="bold")
    ax.text(r,  -0.02, f"+{int(scale_max)}%", ha="right", va="top", fontsize=7.5, color=THEME["muted2"])

    # Needle
    ang = math.radians(angle)
    x = math.cos(ang) * (r - 0.10)
    y = math.sin(ang) * (r - 0.10)
    ax.plot([0, x], [0, y], linewidth=2.2, color=THEME["text"], solid_capstyle="round")
    ax.add_patch(Circle((0, 0), 0.04, facecolor=THEME["text"], edgecolor="none"))

    # Value text
    val_color = THEME["good"] if v > 0 else THEME["bad"] if v < 0 else THEME["muted"]
    sign = "+" if v > 0 else ""
    ax.text(0, 0.42, f"{sign}{v:.1f}%", ha="center", va="center",
            fontsize=13, fontweight="bold", color=val_color)

    # Label
    ax.text(0, 0.17, label, ha="center", va="center",
            fontsize=8.6, color=THEME["text"], fontweight="bold")

    # Subtitle
    ax.text(0, -0.11, f"Échelle ±{int(scale_max)}%", ha="center", va="center",
            fontsize=7.2, color=THEME["muted2"])


def plot_benchmark_speedometers(title: str, items: List[Tuple[str, float, str]], scale_max: Optional[float] = None) -> bytes:
    _mpl_setup()
    vals = [float(v or 0) for _, v, _ in items]
    if scale_max is None:
        scale_max = nice_scale_max(vals)

    fig = plt.figure(figsize=(9.2, 2.9), dpi=180)
    fig.patch.set_facecolor("white")
    fig.suptitle(title, fontsize=11, fontweight="bold", y=0.98, color=THEME["text"])

    n = len(items)
    for i, (lbl, val, col) in enumerate(items):
        ax = fig.add_subplot(1, n, i + 1)
        _draw_speedometer(ax, val, lbl, col, scale_max)

    fig.tight_layout(rect=[0, 0, 1, 0.92])
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════
# 6. HTML EMAIL (Sunêlia-inspired UI)
# ═══════════════════════════════════════════════════════

def build_email_html(D: dict, cids: dict) -> str:
    meta = D.get("meta", {})
    kf = D.get("kpi_fermes", {})
    kt = D.get("kpi_total", {})
    produits = D.get("produits", {})
    saison = D.get("saisonnalite", [])
    bassins = D.get("bassins", [])
    insights = D.get("insights", [])

    date_obs = meta.get("date_observation", "")
    date_comp = meta.get("date_comparaison", "")
    camp = meta.get("camping_name", "Camping Sunêlia")
    yN = yk(date_obs)
    yN1 = yk(date_comp)

    tot_ca = float(produits.get("total_ca", 0) or 0)
    loc_ca = float(produits.get("location", {}).get("ca", 0) or 0)
    emp_ca = float(produits.get("emplacement", {}).get("ca", 0) or 0)
    loc_var = float(produits.get("location", {}).get("var_pct", 0) or 0)
    emp_var = float(produits.get("emplacement", {}).get("var_pct", 0) or 0)
    loc_pct = (loc_ca / tot_ca * 100) if tot_ca else 0.0

    # insight color mapping (premium)
    def insight_style(t: str):
        if t == "positive":
            return ("#e8fbf5", THEME["good"], "#00b089")
        if t == "warning":
            return ("#fff7e6", THEME["warn"], "#f2b640")
        if t == "info":
            return ("#e9f6fb", THEME["info"], "#1c6e8c")
        return ("#ffecec", THEME["bad"], "#d64545")

    html = f"""
<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width">
  <title>Reporting Sunêlia</title>
</head>
<body style="margin:0;padding:0;background:{THEME['bg']};">
  <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;background:{THEME['bg']};">
    <tr>
      <td align="center" style="padding:18px 10px;">

        <table role="presentation" cellpadding="0" cellspacing="0" width="680"
               style="border-collapse:collapse;background:{THEME['surface']};border:1px solid {THEME['border']};border-radius:16px;overflow:hidden;">

          <!-- HEADER (Sunêlia teal gradient) -->
          <tr>
            <td style="padding:22px 22px 18px;
                       background:linear-gradient(135deg,{THEME['teal_500']} 0%, {THEME['teal_700']} 45%, {THEME['teal_900']} 100%);">
              <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;">
                <tr>
                  <td style="color:#ffffff;font-size:12px;font-weight:900;letter-spacing:2px;text-transform:uppercase;">
                    SUNÊLIA VACANCES
                  </td>
                  <td align="right">
                    <span style="display:inline-block;background:rgba(255,255,255,.14);
                                 border:1px solid rgba(255,255,255,.22);color:#ffffff;
                                 font-size:11px;padding:4px 10px;border-radius:999px;">
                      📊 Données au {hesc(date_obs)}
                    </span>
                  </td>
                </tr>
              </table>

              <div style="margin-top:12px;color:#ffffff;font-size:20px;font-weight:900;letter-spacing:.2px;">
                {hesc(camp)}
              </div>
              <div style="margin-top:6px;color:rgba(255,255,255,.78);font-size:12px;line-height:1.4;">
                Vue synthétique — Saison {hesc(yN)} vs {hesc(yN1)} • Réservations fermes hors résidentiel
              </div>

              <div style="margin-top:14px;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.18);
                          border-radius:14px;padding:10px 12px;">
                <div style="font-size:11px;color:rgba(255,255,255,.85);">
                  Comparaison au {hesc(date_comp)} • Montants TTC • Rapport IA (Einstein/Claude)
                </div>
              </div>
            </td>
          </tr>

          {section_title("Indicateurs clés — Réservations fermes")}
          <tr>
            <td style="padding:12px 22px 6px;">
              <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:separate;border-spacing:10px;">
                <tr>
                  {kpi_cell("Chiffre d'affaires TTC",
                            fmt_eur(kf.get("ca",{}).get("n")),
                            f"vs {fmt_eur(kf.get('ca',{}).get('n1'))} en N-1",
                            badge_html(kf.get("ca",{}).get("var_pct"), "pct"))}
                  {kpi_cell("Nombre de séjours",
                            fmt_fr(kf.get("sejours",{}).get("n"),0),
                            f"vs {fmt_fr(kf.get('sejours',{}).get('n1'),0)} en N-1",
                            badge_html(kf.get("sejours",{}).get("var_pct"), "pct"))}
                  {kpi_cell("Nombre de nuits",
                            fmt_fr(kf.get("nuits",{}).get("n"),0),
                            f"vs {fmt_fr(kf.get('nuits',{}).get('n1'),0)} en N-1",
                            badge_html(kf.get("nuits",{}).get("var_pct"), "pct"))}
                </tr>
                <tr>
                  {kpi_cell("Prix moyen / nuit",
                            f"{fmt_fr(kf.get('prix_moyen_nuit',{}).get('n'),2)} €",
                            f"vs {fmt_fr(kf.get('prix_moyen_nuit',{}).get('n1'),2)} € en N-1",
                            badge_html(kf.get("prix_moyen_nuit",{}).get("var_pct"), "pct"))}
                  {kpi_cell("Taux d'occupation",
                            f"{fmt_fr(kf.get('taux_occupation',{}).get('n'),1)}%",
                            f"vs {fmt_fr(kf.get('taux_occupation',{}).get('n1'),1)}% en N-1",
                            badge_html(kf.get("taux_occupation",{}).get("delta_pts"), "pts"))}
                  {kpi_cell("RevPar",
                            fmt_eur(kf.get("revpar",{}).get("n")),
                            f"vs {fmt_eur(kf.get('revpar',{}).get('n1'))} en N-1",
                            badge_html(kf.get("revpar",{}).get("var_pct"), "pct"))}
                </tr>
              </table>
            </td>
          </tr>

          {section_title("Positionnement vs Réseau & Région")}
          {small_note("Jauges premium : 0 au centre, négatif à gauche, positif à droite. Échelle adaptée automatiquement.")}
          <tr>
            <td style="padding:8px 22px 10px;">
              <img src="cid:{hesc(cids['bm_ca'])}" alt="Benchmark CA"
                   style="width:100%;max-width:636px;height:auto;border:1px solid {THEME['border']};
                          border-radius:14px;display:block;background:{THEME['surface']};">
            </td>
          </tr>
          <tr>
            <td style="padding:6px 22px 18px;">
              <img src="cid:{hesc(cids['bm_sj'])}" alt="Benchmark séjours"
                   style="width:100%;max-width:636px;height:auto;border:1px solid {THEME['border']};
                          border-radius:14px;display:block;background:{THEME['surface']};">
            </td>
          </tr>

          {section_title("Montée en charge — CA cumulé")}
          <tr>
            <td style="padding:10px 22px 18px;">
              <img src="cid:{hesc(cids['mc'])}" alt="Montée en charge"
                   style="width:100%;max-width:636px;height:auto;border:1px solid {THEME['border']};
                          border-radius:14px;display:block;background:{THEME['surface']};">
              <div style="font-size:11px;color:{THEME['muted']};margin-top:8px;">
                Courbe générée en image (compatible email) — Saison {hesc(yN)} vs {hesc(yN1)}
              </div>
            </td>
          </tr>

          {section_title("CA par semaine d'occupation")}
          <tr>
            <td style="padding:10px 22px 18px;">
              <img src="cid:{hesc(cids['wk'])}" alt="CA par semaine"
                   style="width:100%;max-width:636px;height:auto;border:1px solid {THEME['border']};
                          border-radius:14px;display:block;background:{THEME['surface']};">
            </td>
          </tr>

          {section_title("Répartition par type de produit")}
          <tr>
            <td style="padding:10px 22px 18px;">
              <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;">
                <tr>
                  <td width="240" style="vertical-align:top;padding-right:12px;">
                    <img src="cid:{hesc(cids['donut'])}" alt="Répartition produits"
                         style="width:100%;max-width:240px;height:auto;border:1px solid {THEME['border']};
                                border-radius:14px;display:block;background:{THEME['surface']};">
                  </td>
                  <td style="vertical-align:top;">
                    <div style="font-size:12px;color:{THEME['text']};margin-bottom:10px;">
                      <strong>Total hors résidentiel :</strong> {hesc(fmt_eur(tot_ca))}
                    </div>
                    <table role="presentation" cellpadding="0" cellspacing="0" width="100%"
                           style="border-collapse:collapse;font-size:12px;background:{THEME['panel']};
                                  border:1px solid {THEME['border']};border-radius:14px;overflow:hidden;">
                      <tr>
                        <td style="padding:10px 12px;">
                          <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:{THEME['teal_500']};margin-right:8px;"></span>
                          Location
                        </td>
                        <td align="right" style="padding:10px 12px;font-weight:900;color:{THEME['text']};">
                          {hesc(fmt_eur(loc_ca))}
                          <span style="font-size:11px;color:{THEME['good'] if loc_var>=0 else THEME['bad']};">({hesc(fmt_pct(loc_var))})</span>
                        </td>
                      </tr>
                      <tr>
                        <td style="padding:10px 12px;border-top:1px solid {THEME['border']};">
                          <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:{THEME['sun']};margin-right:8px;"></span>
                          Emplacements
                        </td>
                        <td align="right" style="padding:10px 12px;border-top:1px solid {THEME['border']};font-weight:900;color:{THEME['text']};">
                          {hesc(fmt_eur(emp_ca))}
                          <span style="font-size:11px;color:{THEME['good'] if emp_var>=0 else THEME['bad']};">({hesc(fmt_pct(emp_var))})</span>
                        </td>
                      </tr>
                    </table>

                    <div style="margin-top:10px;font-size:11px;color:{THEME['muted']};">
                      La location représente <strong style="color:{THEME['text']};">{hesc(fmt_fr(loc_pct,1))}%</strong> du CA.
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          {section_title("Activité par saisonnalité")}
          <tr>
            <td style="padding:10px 22px 18px;">
              <table role="presentation" cellpadding="0" cellspacing="0" width="100%"
                     style="border-collapse:collapse;font-size:12px;border:1px solid {THEME['border']};border-radius:14px;overflow:hidden;">
                <tr>
                  <th style="background:{THEME['teal_900']};color:#fff;padding:10px 10px;text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:.6px;">Période</th>
                  <th style="background:{THEME['teal_900']};color:#fff;padding:10px 10px;text-align:right;font-size:10px;text-transform:uppercase;letter-spacing:.6px;">CA N</th>
                  <th style="background:{THEME['teal_900']};color:#fff;padding:10px 10px;text-align:right;font-size:10px;text-transform:uppercase;letter-spacing:.6px;">CA N-1</th>
                  <th style="background:{THEME['teal_900']};color:#fff;padding:10px 10px;text-align:right;font-size:10px;text-transform:uppercase;letter-spacing:.6px;">Var %</th>
                  <th style="background:{THEME['teal_900']};color:#fff;padding:10px 10px;text-align:right;font-size:10px;text-transform:uppercase;letter-spacing:.6px;">Séjours</th>
                  <th style="background:{THEME['teal_900']};color:#fff;padding:10px 10px;text-align:right;font-size:10px;text-transform:uppercase;letter-spacing:.6px;">Taux occ.</th>
                </tr>
                {"".join([
                    f"<tr style=\"background:{'#eafaf6' if s.get('is_haute') else ('#ffffff' if i%2==0 else THEME['panel'])};\">"
                    f"<td style=\"padding:9px 10px;border-bottom:1px solid {THEME['border']};color:{THEME['text']};\">"
                    f"{'🌞 ' if s.get('is_haute') else ''}{hesc(str(s.get('periode','')))}</td>"
                    f"<td style=\"padding:9px 10px;border-bottom:1px solid {THEME['border']};text-align:right;color:{THEME['text']};\">{hesc(fmt_eur(s.get('ca_n')))}</td>"
                    f"<td style=\"padding:9px 10px;border-bottom:1px solid {THEME['border']};text-align:right;color:{THEME['text']};\">{hesc(fmt_eur(s.get('ca_n1')))}</td>"
                    f"<td style=\"padding:9px 10px;border-bottom:1px solid {THEME['border']};text-align:right;font-weight:900;color:{THEME['good'] if float(s.get('var_pct') or 0)>=0 else THEME['bad']};\">{hesc(fmt_pct(s.get('var_pct')))}</td>"
                    f"<td style=\"padding:9px 10px;border-bottom:1px solid {THEME['border']};text-align:right;color:{THEME['text']};\">{hesc(fmt_fr(s.get('sejours'),0))}</td>"
                    f"<td style=\"padding:9px 10px;border-bottom:1px solid {THEME['border']};text-align:right;color:{THEME['text']};\">{hesc(fmt_fr(s.get('taux_occ'),1))}%</td>"
                    f"</tr>"
                    for i, s in enumerate(saison)
                ])}
              </table>
            </td>
          </tr>

          {section_title("Top bassins émetteurs")}
          <tr>
            <td style="padding:10px 22px 18px;">
              <table role="presentation" cellpadding="0" cellspacing="0" width="100%"
                     style="border-collapse:collapse;font-size:12px;border:1px solid {THEME['border']};border-radius:14px;overflow:hidden;">
                <tr>
                  <th style="background:{THEME['teal_900']};color:#fff;padding:10px 10px;text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:.6px;">Région</th>
                  <th style="background:{THEME['teal_900']};color:#fff;padding:10px 10px;text-align:right;font-size:10px;text-transform:uppercase;letter-spacing:.6px;">CA N</th>
                  <th style="background:{THEME['teal_900']};color:#fff;padding:10px 10px;text-align:right;font-size:10px;text-transform:uppercase;letter-spacing:.6px;">% total</th>
                  <th style="background:{THEME['teal_900']};color:#fff;padding:10px 10px;text-align:right;font-size:10px;text-transform:uppercase;letter-spacing:.6px;">Évolution</th>
                </tr>
                {"".join([
                    f"<tr style=\"background:{'#ffffff' if i%2==0 else THEME['panel']};\">"
                    f"<td style=\"padding:9px 10px;border-bottom:1px solid {THEME['border']};color:{THEME['text']};\">{hesc(('🥇 ' if i==0 else '🥈 ' if i==1 else '🥉 ' if i==2 else '') + str(b.get('region','')))}</td>"
                    f"<td style=\"padding:9px 10px;border-bottom:1px solid {THEME['border']};text-align:right;color:{THEME['text']};\">{hesc(fmt_eur(b.get('ca')))}</td>"
                    f"<td style=\"padding:9px 10px;border-bottom:1px solid {THEME['border']};text-align:right;color:{THEME['text']};\">{hesc(fmt_fr(b.get('pct_total'),1))}%</td>"
                    f"<td style=\"padding:9px 10px;border-bottom:1px solid {THEME['border']};text-align:right;font-weight:900;color:{THEME['good'] if float(b.get('var_pct') or 0)>=0 else THEME['bad']};\">{hesc(fmt_pct(b.get('var_pct')))}</td>"
                    f"</tr>"
                    for i, b in enumerate(bassins)
                ])}
              </table>
            </td>
          </tr>

          {section_title("Analyse & points d'attention")}
          <tr>
            <td style="padding:10px 22px 18px;">
              <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:separate;border-spacing:10px;">
                {"".join([
                    "<tr>" +
                    "".join([
                        (lambda bg, fg, bd:
                            f"<td width=\"50%\" style=\"vertical-align:top;\">"
                            f"<div style=\"padding:14px 14px;border-radius:14px;border:1px solid {THEME['border']};"
                            f"border-left:5px solid {bd};background:{bg};\">"
                            f"<div style=\"font-size:11px;font-weight:900;text-transform:uppercase;letter-spacing:.6px;"
                            f"color:{fg};margin-bottom:6px;\">{hesc(str(it.get('title','')))}</div>"
                            f"<div style=\"font-size:12px;color:{THEME['text']};line-height:1.45;\">{hesc(str(it.get('text','')))}</div>"
                            f"</div></td>"
                        )(*insight_style(it.get('type','info')))
                        for it in insights[i:i+2]
                    ]) +
                    "</tr>"
                    for i in range(0, len(insights), 2)
                ])}
              </table>
            </td>
          </tr>

          {section_title("Synthèse Options + Réservations")}
          <tr>
            <td style="padding:12px 22px 22px;">
              <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:separate;border-spacing:10px;">
                <tr>
                  {kpi_cell("CA Total",
                            fmt_eur(kt.get("ca",{}).get("n")),
                            "Options + réservations (hors résidentiel)",
                            badge_html(kt.get("ca",{}).get("var_pct"), "pct"))}
                  {kpi_cell("Séjours totaux",
                            fmt_fr(kt.get("sejours",{}).get("n"),0),
                            "Options + réservations (hors résidentiel)",
                            badge_html(kt.get("sejours",{}).get("var_pct"), "pct"))}
                  <td style="vertical-align:top;border:1px solid {THEME['border']};padding:12px 12px;border-radius:12px;background:{THEME['surface']};">
                    <div style="font-size:10px;text-transform:uppercase;letter-spacing:1px;color:{THEME['muted2']};font-weight:800;margin-bottom:6px;">Stock</div>
                    <div style="font-size:20px;font-weight:900;color:{THEME['text']};line-height:1.1;margin-bottom:4px;">{hesc(str(kt.get("stock","-")))}</div>
                    <div style="font-size:11px;color:{THEME['muted']};margin-bottom:8px;">Capacité / stock</div>
                    <span style="display:inline-block;font-size:11px;font-weight:800;padding:3px 9px;border-radius:999px;background:#edf3f4;color:{THEME['muted']};">= stable</span>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- FOOTER (Sunêlia teal) -->
          <tr>
            <td style="padding:16px 22px;text-align:center;
                       background:linear-gradient(135deg,{THEME['teal_900']} 0%, {THEME['teal_700']} 50%, {THEME['teal_500']} 100%);">
              <div style="color:#ffffff;font-weight:900;font-size:12px;letter-spacing:2px;text-transform:uppercase;">
                Sunêlia Vacances
              </div>
              <div style="color:rgba(255,255,255,.75);font-size:11px;margin-top:6px;">
                Rapport généré par IA — Données CRM au {hesc(date_obs)} comparées au {hesc(date_comp)}
              </div>
              <div style="color:rgba(255,255,255,.65);font-size:11px;margin-top:4px;">
                {hesc(camp)} • Montants TTC
              </div>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</body>
</html>
    """.strip()

    return html


def build_plain_text(D: dict) -> str:
    meta = D.get("meta", {})
    camp = meta.get("camping_name", "Camping Sunêlia")
    date_obs = meta.get("date_observation", "")
    date_comp = meta.get("date_comparaison", "")
    kf = D.get("kpi_fermes", {})

    return "\n".join([
        f"Reporting Sunêlia - {camp}",
        f"Données au {date_obs} (comparaison {date_comp})",
        "",
        "Réservations fermes (hors résidentiel):",
        f"- CA: {fmt_eur(kf.get('ca',{}).get('n'))} (var {fmt_pct(kf.get('ca',{}).get('var_pct'))})",
        f"- Séjours: {fmt_fr(kf.get('sejours',{}).get('n'),0)} (var {fmt_pct(kf.get('sejours',{}).get('var_pct'))})",
        f"- Nuits: {fmt_fr(kf.get('nuits',{}).get('n'),0)} (var {fmt_pct(kf.get('nuits',{}).get('var_pct'))})",
        f"- PMN: {fmt_fr(kf.get('prix_moyen_nuit',{}).get('n'),2)} € (var {fmt_pct(kf.get('prix_moyen_nuit',{}).get('var_pct'))})",
        f"- TO: {fmt_fr(kf.get('taux_occupation',{}).get('n'),1)}% (delta {fmt_pts(kf.get('taux_occupation',{}).get('delta_pts'))})",
        f"- RevPar: {fmt_eur(kf.get('revpar',{}).get('n'))} (var {fmt_pct(kf.get('revpar',{}).get('var_pct'))})",
        "",
        "NB: Les graphiques sont inclus en images dans la version HTML du mail."
    ])


def build_eml_message(
    D: dict,
    from_addr: str,
    to_addr: str,
    subject: str,
    img_mc: bytes,
    img_wk: bytes,
    img_donut: bytes,
    img_bm_ca: bytes,
    img_bm_sj: bytes
) -> bytes:
    cids = {"mc": "mc_chart", "wk": "wk_chart", "donut": "donut_chart", "bm_ca": "bm_ca", "bm_sj": "bm_sj"}

    msg_root = MIMEMultipart("related")
    msg_root["Subject"] = subject
    msg_root["From"] = from_addr
    msg_root["To"] = to_addr

    msg_alt = MIMEMultipart("alternative")
    msg_root.attach(msg_alt)

    txt = build_plain_text(D)
    html = build_email_html(D, cids)

    msg_alt.attach(MIMEText(txt, "plain", "utf-8"))
    msg_alt.attach(MIMEText(html, "html", "utf-8"))

    def attach_img(img_bytes: bytes, cid: str, filename: str):
        part = MIMEImage(img_bytes, _subtype="png")
        part.add_header("Content-ID", f"<{cid}>")
        part.add_header("Content-Disposition", "inline", filename=filename)
        msg_root.attach(part)

    attach_img(img_mc, cids["mc"], "montee_charge.png")
    attach_img(img_wk, cids["wk"], "ventes_semaine.png")
    attach_img(img_donut, cids["donut"], "produits.png")
    attach_img(img_bm_ca, cids["bm_ca"], "benchmark_ca.png")
    attach_img(img_bm_sj, cids["bm_sj"], "benchmark_sejours.png")

    return msg_root.as_bytes(policy=policy.SMTP)


# ═══════════════════════════════════════════════════════
# 7. MAIN
# ═══════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description="Génère un reporting email (.eml) Sunêlia via IA — UI thème Sunêlia + benchmark premium")
    parser.add_argument("excel", help="Chemin vers le fichier Excel (.xlsx)")
    parser.add_argument("--org", default="PROD", help="Org Salesforce (défaut: PROD)")
    parser.add_argument("--output", default="", help="Dossier de sortie (défaut: même que le fichier)")
    parser.add_argument("--dry-run", action="store_true", help="Sauvegarde le prompt sans appeler l'API")
    parser.add_argument("--from", dest="from_addr", default="reporting@sunelia.local", help="Adresse expéditeur")
    parser.add_argument("--to", dest="to_addr", default="destinataire@exemple.com", help="Adresse destinataire")
    parser.add_argument("--subject", default="", help="Sujet du mail (défaut auto)")
    args = parser.parse_args()

    excel_path = Path(args.excel)
    if not excel_path.exists():
        print(f"❌ Fichier introuvable: {excel_path}")
        sys.exit(1)

    output_dir = Path(args.output) if args.output else excel_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    print()
    print("  ╔══════════════════════════════════════════════════╗")
    print("  ║   SUNÊLIA — Génération de Reporting IA (EMAIL)  ║")
    print("  ╚══════════════════════════════════════════════════╝")
    print()

    # 1) Extraction
    print("  [1/4] Extraction des données Excel...")
    data_text = extract_excel(str(excel_path))
    print(f"        ✓ {len(data_text) / 1024:.1f} Ko de données extraites")

    # 2) Prompt
    print("  [2/4] Construction du prompt...")
    prompt = build_prompt(data_text)
    print(f"        ✓ Prompt: {len(prompt)} chars (~{len(prompt)//4} tokens)")

    if args.dry_run:
        debug_path = output_dir / "prompt_debug.txt"
        debug_path.write_text(prompt, encoding="utf-8")
        print(f"\n  [DRY RUN] Prompt sauvegardé: {debug_path}")
        sys.exit(0)

    # 3) IA
    print("  [3/4] Authentification Salesforce...")
    token, instance_url = get_sf_auth(args.org)
    print("        ✓ Token obtenu")

    print("  [3/4] Appel IA en cours...")
    t0 = time.time()
    raw_response = call_einstein(prompt, token, instance_url)
    elapsed = time.time() - t0
    print(f"        ✓ Réponse reçue en {elapsed:.1f}s ({len(raw_response)} chars)")

    debug_raw = output_dir / "ai_response_raw.txt"
    debug_raw.write_text(raw_response, encoding="utf-8")

    # 4) JSON + EML
    print("  [4/4] Génération du mail .eml (HTML + images inline)...")
    try:
        json_str = clean_json(raw_response)
    except ValueError:
        debug_path = output_dir / "ai_response_debug.txt"
        debug_path.write_text(raw_response, encoding="utf-8")
        print(f"  ❌ JSON invalide. Réponse brute sauvegardée: {debug_path}")
        sys.exit(1)

    debug_json = output_dir / "report_data.json"
    debug_json.write_text(json_str, encoding="utf-8")
    print(f"        ✓ JSON nettoyé sauvegardé: {debug_json.name}")

    D = json.loads(json_str)
    print(f"        ✓ JSON valide ({len(D.get('montee_charge',[]))} points montée en charge)")
    print(f"        ✓ Camping: {D.get('meta',{}).get('camping_name','?')}")

    # Graphiques
    img_mc = plot_montee_charge(D)
    img_wk = plot_ventes_semaine(D)
    img_donut = plot_produits_donut(D)

    # PREMIUM Benchmark speedometers (3 jauges horizontales)
    bm = D.get("benchmark", {}) or {}
    camp_name_short = D.get("meta", {}).get("camping_name", "Camping").replace("Camping Sunêlia ", "")
    region_lbl = f"{bm.get('region_label','Région')} ({bm.get('region_nb','')})".strip()
    reseau_lbl = f"{bm.get('reseau_label','Réseau')} ({bm.get('reseau_nb','')})".strip()

    ca_vals = bm.get("ca", {}) or {}
    sj_vals = bm.get("sejours", {}) or {}

    # Couleurs différenciées mais cohérentes avec le thème
    c_camping = THEME["teal_500"]
    c_region  = THEME["teal_900"]
    c_reseau  = THEME["teal_600"]

    img_bm_ca = plot_benchmark_speedometers(
        "Positionnement — Variation du CA",
        [
            (camp_name_short, float(ca_vals.get("camping", 0) or 0), c_camping),
            (region_lbl,      float(ca_vals.get("region", 0) or 0), c_region),
            (reseau_lbl,      float(ca_vals.get("reseau", 0) or 0), c_reseau),
        ],
        scale_max=None
    )

    img_bm_sj = plot_benchmark_speedometers(
        "Positionnement — Variation du nombre de séjours",
        [
            (camp_name_short, float(sj_vals.get("camping", 0) or 0), c_camping),
            (region_lbl,      float(sj_vals.get("region", 0) or 0), c_region),
            (reseau_lbl,      float(sj_vals.get("reseau", 0) or 0), c_reseau),
        ],
        scale_max=None
    )

    if not args.subject:
        camp = D.get("meta", {}).get("camping_name", "Camping Sunêlia")
        date_obs = D.get("meta", {}).get("date_observation", "")
        args.subject = f"Reporting Sunêlia — {camp} — {date_obs}"

    eml_bytes = build_eml_message(
        D,
        from_addr=args.from_addr,
        to_addr=args.to_addr,
        subject=args.subject,
        img_mc=img_mc,
        img_wk=img_wk,
        img_donut=img_donut,
        img_bm_ca=img_bm_ca,
        img_bm_sj=img_bm_sj
    )

    base_name = excel_path.stem.replace("Sunelia_Rapports_indiv_pour_groupe_", "")
    output_file = output_dir / f"Reporting_{base_name}.eml"
    output_file.write_bytes(eml_bytes)

    print()
    print("  ╔══════════════════════════════════════════════════╗")
    print("  ║   ✓ MAIL .EML GÉNÉRÉ AVEC SUCCÈS                 ║")
    print("  ╚══════════════════════════════════════════════════╝")
    print(f"  → {output_file}")
    print("  (Ouvre le .eml dans Outlook/Apple Mail/Thunderbird, ou importe-le dans Gmail.)")
    print()


if __name__ == "__main__":
    main()