import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from fpdf import FPDF
from pypdf import PdfReader
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime, timezone
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
import tempfile
import base64
import gzip
import json
import io
import os
import re

try:
    from rapidocr_onnxruntime import RapidOCR
    OCR_AVAILABLE = True
except Exception:
    OCR_AVAILABLE = False

# --- SAYFA AYARLARI ---
st.set_page_config(
    page_title="TR-CBAM | AB Sınırda Karbon (SKDM) Ceza Kalkanı",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- 10 HAZIR KURUMSAL RENK PALETİ (SAP FIORI QUARTZ / HORIZON & DARK SLATE) ---
FIORI_THEMES = {
    "1. SAP Fiori Quartz (Varsayılan Kurumsal)": {
        "is_dark": False,
        "shell_bg": "#354A5F", "shell_text": "#FFFFFF",
        "bg_canvas": "#F5F6F7", "bg_card": "#FFFFFF",
        "text_primary": "#1D2D3E", "text_secondary": "#556B82",
        "accent": "#0A6ED1", "accent_text": "#FFFFFF",
        "danger": "#BB0000", "danger_bg": "rgba(187, 0, 0, 0.08)",
        "success": "#107E3E", "success_bg": "rgba(16, 126, 62, 0.09)",
        "border": "#D9E0E6"
    },
    "2. SAP Horizon Morning": {
        "is_dark": False,
        "shell_bg": "#1A2733", "shell_text": "#FFFFFF",
        "bg_canvas": "#EFF3F7", "bg_card": "#FFFFFF",
        "text_primary": "#131E29", "text_secondary": "#5B738B",
        "accent": "#0070F2", "accent_text": "#FFFFFF",
        "danger": "#D20A0A", "danger_bg": "rgba(210, 10, 10, 0.08)",
        "success": "#188918", "success_bg": "rgba(24, 137, 24, 0.09)",
        "border": "#D5DFE8"
    },
    "3. Sıcak Kireçtaşı (Warm Limestone)": {
        "is_dark": False,
        "shell_bg": "#292524", "shell_text": "#FAF8F5",
        "bg_canvas": "#F7F5F0", "bg_card": "#FCFBF9",
        "text_primary": "#1C1917", "text_secondary": "#57534E",
        "accent": "#B45309", "accent_text": "#FFFFFF",
        "danger": "#C2410C", "danger_bg": "rgba(194, 65, 12, 0.08)",
        "success": "#15803D", "success_bg": "rgba(21, 128, 61, 0.09)",
        "border": "#E7E2D8"
    },
    "4. İsviçre Bankacılığı (Swiss Vault)": {
        "is_dark": False,
        "shell_bg": "#111827", "shell_text": "#F9FAFB",
        "bg_canvas": "#F3F4F6", "bg_card": "#FFFFFF",
        "text_primary": "#111827", "text_secondary": "#4B5563",
        "accent": "#1F2937", "accent_text": "#FFFFFF",
        "danger": "#DC2626", "danger_bg": "rgba(220, 38, 38, 0.08)",
        "success": "#047857", "success_bg": "rgba(4, 120, 87, 0.09)",
        "border": "#D1D5DB"
    },
    "5. Yeşil Finans (ESG Sage)": {
        "is_dark": False,
        "shell_bg": "#132A22", "shell_text": "#ECFDF5",
        "bg_canvas": "#F2F5F3", "bg_card": "#FAFDFB",
        "text_primary": "#142921", "text_secondary": "#4B6358",
        "accent": "#1B4332", "accent_text": "#FFFFFF",
        "danger": "#B91C1C", "danger_bg": "rgba(185, 28, 28, 0.08)",
        "success": "#16A34A", "success_bg": "rgba(22, 163, 74, 0.09)",
        "border": "#D4E0D9"
    },
    "6. Gece Arduvazı (Midnight Slate Dark)": {
        "is_dark": True,
        "shell_bg": "#080D1A", "shell_text": "#F1F5F9",
        "bg_canvas": "#0B1120", "bg_card": "#131C31",
        "text_primary": "#F1F5F9", "text_secondary": "#94A3B8",
        "accent": "#3B82F6", "accent_text": "#FFFFFF",
        "danger": "#F87171", "danger_bg": "rgba(248, 113, 113, 0.12)",
        "success": "#34D399", "success_bg": "rgba(52, 211, 153, 0.12)",
        "border": "#24314D"
    },
    "7. SAP Fiori Evening (Dark Horizon)": {
        "is_dark": True,
        "shell_bg": "#12171C", "shell_text": "#EAECEE",
        "bg_canvas": "#1C2228", "bg_card": "#232A31",
        "text_primary": "#EAECEE", "text_secondary": "#89919A",
        "accent": "#4DB1FF", "accent_text": "#0B1929",
        "danger": "#FF8888", "danger_bg": "rgba(255, 136, 136, 0.12)",
        "success": "#5DC122", "success_bg": "rgba(93, 193, 34, 0.12)",
        "border": "#323C46"
    },
    "8. Endüstriyel Karbon (Graphite Steel)": {
        "is_dark": True,
        "shell_bg": "#0D0E10", "shell_text": "#E4E4E7",
        "bg_canvas": "#121316", "bg_card": "#1A1C20",
        "text_primary": "#E4E4E7", "text_secondary": "#A1A1AA",
        "accent": "#60A5FA", "accent_text": "#091322",
        "danger": "#EF4444", "danger_bg": "rgba(239, 68, 68, 0.12)",
        "success": "#10B981", "success_bg": "rgba(16, 185, 129, 0.12)",
        "border": "#2A2E35"
    },
    "9. Derin Prusya Mavisi (Deep Prussian)": {
        "is_dark": True,
        "shell_bg": "#071019", "shell_text": "#E2E8F0",
        "bg_canvas": "#0A1622", "bg_card": "#112233",
        "text_primary": "#E2E8F0", "text_secondary": "#8FA3B8",
        "accent": "#38BDF8", "accent_text": "#061826",
        "danger": "#FB7185", "danger_bg": "rgba(251, 113, 133, 0.12)",
        "success": "#2DD4BF", "success_bg": "rgba(45, 212, 191, 0.12)",
        "border": "#1E3852"
    },
    "10. Obsidyen & Kehribar (Private Banking Dark)": {
        "is_dark": True,
        "shell_bg": "#0E0D0B", "shell_text": "#F5F5F4",
        "bg_canvas": "#141210", "bg_card": "#1E1B18",
        "text_primary": "#F5F5F4", "text_secondary": "#A8A29E",
        "accent": "#D97706", "accent_text": "#FFFFFF",
        "danger": "#F87171", "danger_bg": "rgba(248, 113, 113, 0.12)",
        "success": "#4ADE80", "success_bg": "rgba(74, 222, 128, 0.12)",
        "border": "#332E29"
    },
}

# --- 10 KURUMSAL GOOGLE FONTS YAZI TİPİ ---
FIORI_FONTS = {
    "IBM Plex Sans": "IBM+Plex+Sans:wght@400;500;600;700",
    "Plus Jakarta Sans": "Plus+Jakarta+Sans:wght@400;500;600;700",
    "Inter": "Inter:wght@400;500;600;700",
    "Manrope": "Manrope:wght@400;500;600;700",
    "DM Sans": "DM+Sans:wght@400;500;600;700",
    "Public Sans": "Public+Sans:wght@400;500;600;700",
    "Work Sans": "Work+Sans:wght@400;500;600;700",
    "Source Sans 3": "Source+Sans+3:wght@400;600;700",
    "Space Grotesk": "Space+Grotesk:wght@400;500;600;700",
    "JetBrains Mono": "JetBrains+Mono:wght@400;500;600;700",
}

def inject_fiori_design_system(theme_name: str, font_name: str):
    t = FIORI_THEMES[theme_name]
    font_query = FIORI_FONTS[font_name]
    input_bg = "#0F172A" if t["is_dark"] else "#FFFFFF"
    input_text = "#F1F5F9" if t["is_dark"] else "#1D2D3E"

    st.markdown(f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family={font_query}&family=JetBrains+Mono:wght@500;600;700&display=swap');

        :root {{
            --shell-bg: {t['shell_bg']};
            --shell-text: {t['shell_text']};
            --bg-canvas: {t['bg_canvas']};
            --bg-card: {t['bg_card']};
            --text-primary: {t['text_primary']};
            --text-secondary: {t['text_secondary']};
            --accent: {t['accent']};
            --accent-text: {t['accent_text']};
            --danger: {t['danger']};
            --danger-bg: {t['danger_bg']};
            --success: {t['success']};
            --success-bg: {t['success_bg']};
            --border: {t['border']};
        }}

        /* 1. Ham Streamlit İzlerini ve Sol Sidebar'ı Gizle */
        #MainMenu, footer, header[data-testid="stHeader"] {{ visibility: hidden; height: 0; }}
        .stDeployButton, [data-testid="stToolbar"], [data-testid="collapsedControl"] {{ display: none !important; }}
        section[data-testid="stSidebar"] {{ display: none !important; }}

        /* 2. Ana Kanvas & Dinamik Tipografi */
        .stApp {{
            background-color: var(--bg-canvas) !important;
            color: var(--text-primary) !important;
            font-family: '{font_name}', -apple-system, BlinkMacSystemFont, sans-serif !important;
        }}
        .block-container {{
            padding-top: 0.9rem !important;
            padding-bottom: 2.5rem !important;
            max-width: 1460px;
        }}
        h1, h2, h3, h4, p, label, span, div {{
            font-family: '{font_name}', -apple-system, BlinkMacSystemFont, sans-serif;
        }}
        .stMarkdown, .stText, label {{
            color: var(--text-primary) !important;
        }}

        /* Form Input ve Selectbox Kontrast Güvencesi */
        input, textarea, select, [data-baseweb="select"] > div, [data-baseweb="input"] > div {{
            background-color: {input_bg} !important;
            color: {input_text} !important;
            border-color: var(--border) !important;
            border-radius: 6px !important;
        }}

        /* 3. SAP Fiori Shell Bar (Üst Kabuk Çubuğu) */
        .fiori-shell-bar {{
            background-color: var(--shell-bg);
            color: var(--shell-text);
            border-radius: 8px;
            padding: 12px 20px;
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.12);
        }}
        .fiori-brand {{
            display: flex;
            align-items: center;
            gap: 12px;
        }}
        .fiori-logo-box {{
            background-color: var(--accent);
            color: var(--accent-text);
            font-weight: 700;
            font-size: 0.95rem;
            padding: 5px 10px;
            border-radius: 6px;
            letter-spacing: 0.03em;
        }}
        .fiori-title {{
            font-size: 1.12rem;
            font-weight: 700;
            color: var(--shell-text);
        }}
        .fiori-subtitle {{
            font-size: 0.78rem;
            opacity: 0.85;
            color: var(--shell-text);
        }}
        .fiori-compliance-badge {{
            background-color: rgba(16, 185, 129, 0.18);
            color: #34D399;
            border: 1px solid rgba(52, 211, 153, 0.45);
            border-radius: 999px;
            padding: 4px 12px;
            font-size: 0.75rem;
            font-weight: 600;
            white-space: nowrap;
        }}

        /* 4. Fiori Yatay Sekme Navigasyonu (Aktif Sekme Altı 3px Çizgi) */
        .stTabs [data-baseweb="tab-list"] {{
            background-color: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 4px 10px;
            gap: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        }}
        .stTabs [data-baseweb="tab"] {{
            background-color: transparent !important;
            color: var(--text-secondary) !important;
            font-weight: 600 !important;
            font-size: 0.86rem !important;
            padding: 10px 14px !important;
            border-bottom: 3px solid transparent !important;
            border-radius: 0px !important;
        }}
        .stTabs [aria-selected="true"] {{
            color: var(--accent) !important;
            border-bottom: 3px solid var(--accent) !important;
        }}

        /* 5. Modüler Fiori Tile / KPI Kartları */
        .fiori-tile {{
            background-color: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 18px 20px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.05);
            height: 100%;
        }}
        .fiori-tile-danger {{
            border-top: 3px solid var(--danger);
            background: linear-gradient(180deg, var(--danger-bg) 0%, var(--bg-card) 45%);
        }}
        .fiori-tile-neutral {{
            border-top: 3px solid var(--accent);
        }}
        .fiori-tile-success {{
            border-top: 3px solid var(--success);
            background: linear-gradient(180deg, var(--success-bg) 0%, var(--bg-card) 55%);
        }}
        .tile-header {{
            font-size: 0.80rem;
            font-weight: 600;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 0.04em;
            margin-bottom: 10px;
        }}
        .tile-bottom-row {{
            display: flex;
            align-items: flex-end;
            justify-content: space-between;
            gap: 8px;
        }}
        .tile-kpi {{
            font-family: 'JetBrains Mono', monospace !important;
            font-variant-numeric: tabular-nums !important;
            font-size: 1.72rem;
            font-weight: 700;
            line-height: 1.15;
        }}
        .tile-pill {{
            font-size: 0.72rem;
            font-weight: 600;
            padding: 3px 8px;
            border-radius: 6px;
            white-space: nowrap;
        }}
        .tile-footnote {{
            font-size: 0.78rem;
            color: var(--text-secondary);
            margin-top: 8px;
        }}

        /* 6. Butonlar ve Paywall Kartı */
        .stButton > button[kind="primary"], .stFormSubmitButton > button {{
            background-color: var(--accent) !important;
            color: var(--accent-text) !important;
            border: none !important;
            font-weight: 600 !important;
            border-radius: 6px !important;
            padding: 0.55rem 1rem !important;
        }}
        .fiori-paywall {{
            background-color: var(--bg-card);
            border: 1px solid var(--border);
            border-left: 4px solid var(--accent);
            border-radius: 8px;
            padding: 16px 20px;
            margin-top: 12px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.05);
        }}
        .blur-doc-preview {{
            background-color: var(--bg-canvas);
            border: 1px dashed var(--border);
            border-radius: 6px;
            padding: 12px 14px;
            margin-top: 10px;
            font-family: 'JetBrains Mono', monospace !important;
            font-variant-numeric: tabular-nums !important;
            font-size: 0.79rem;
            color: var(--text-primary);
        }}
        .blur-overlay {{
            filter: blur(3.5px);
            user-select: none;
            opacity: 0.6;
        }}
    </style>
    """, unsafe_allow_html=True)

# --- LİSANS ANAHTARI DOĞRULAMA ---
def check_license_active(key_str: str) -> tuple[bool, str]:
    k = (key_str or "").strip().upper()
    if k.startswith("TRCBAM-ENT-"):
        return True, "Kurumsal Danışmanlık / Enterprise Lisans Aktif"
    if k.startswith("TRCBAM-OSB-"):
        return True, "OSB / İhracatçı Birliği Lisansı Aktif"
    if k.startswith("TRCBAM-PRO-") and len(k) >= 15:
        return True, "KOBİ İhracatçı Pro Lisans Aktif"
    return False, "Deneme Modu: Filigranlı Ön İzleme"

# --- SABİT MEVZUAT KATSAYILARI ((EU) 2025/2547 Ek II Bölüm G & ETKB) ---
FUEL_FACTORS = {
    "dogalgaz": {"name": "Doğalgaz (m³)", "ef": 0.00192, "gj_per_unit": 0.0342, "unit": "m³"},
    "tas_komuru": {"name": "Taş Kömürü (Ton)", "ef": 2.420, "gj_per_unit": 25.8, "unit": "Ton"},
    "kok_komuru": {"name": "Kok Kömürü (Ton)", "ef": 3.100, "gj_per_unit": 28.2, "unit": "Ton"},
    "linyit": {"name": "Linyit Kömürü (Ton)", "ef": 1.150, "gj_per_unit": 11.9, "unit": "Ton"},
    "fuel_oil": {"name": "Fuel-Oil (Ton)", "ef": 3.120, "gj_per_unit": 40.4, "unit": "Ton"},
    "lpg": {"name": "LPG (Ton)", "ef": 2.980, "gj_per_unit": 47.3, "unit": "Ton"},
}

EU_DEFAULT_TR_GRID_EF = 0.718

PPA_GRID_OPTIONS = {
    "ETKB Dağıtım Hattı / OSB Bağlantısı (0,469 tCO₂e/MWh)": 0.469,
    "ETKB İletim Hattı / Yüksek Gerilim (0,436 tCO₂e/MWh)": 0.436,
    "Sıfır Emisyonlu Yenilenebilir PPA / GES (0,000 tCO₂e/MWh)": 0.000,
}

ROUTE_LABELS = {
    "(A)": "(A) Gri Klinker / Gri Çimento",
    "(B)": "(B) Beyaz Klinker / Beyaz Çimento",
    "(C)": "(C) Karbon Çelik — Yüksek Fırın (BF/BOF)",
    "(D)": "(D) Karbon Çelik — DRI / EAF",
    "(E)": "(E) Karbon Çelik — Hurda / Elektrik Ark Ocağı (Scrap/EAF)",
    "(F)": "(F) Düşük Alaşımlı Çelik — Yüksek Fırın (BF/BOF)",
    "(G)": "(G) Düşük Alaşımlı Çelik — DRI / EAF",
    "(H)": "(H) Düşük Alaşımlı Çelik — Hurda / Elektrik Ark Ocağı (Scrap/EAF)",
    "(J)": "(J) Yüksek Alaşımlı Çelik — EAF",
    "(K)": "(K) Birincil Alüminyum (Primary Aluminium)",
    "(L)": "(L) İkincil Alüminyum (Secondary / Geri Dönüşüm)",
    "GENEL": "Standart Üretim Rotası (Rota Bağımsız)"
}

YEAR_RULES = {
    "2026 (Kesin Dönem 1. Yıl | Ceza: +%10 | CBAM Faktörü: %97,5)": {
        "std_markup": 1.10, "fert_markup": 1.01, "cbam_factor": 0.975, "year": 2026, "period_code": "(1)"
    },
    "2027 (Kesin Dönem 2. Yıl | Ceza: +%20 | CBAM Faktörü: %95,0)": {
        "std_markup": 1.20, "fert_markup": 1.01, "cbam_factor": 0.950, "year": 2027, "period_code": "(1)"
    },
    "2028 (Kesin Dönem 3. Yıl | Ceza: +%30 | CBAM Faktörü: %90,0)": {
        "std_markup": 1.30, "fert_markup": 1.01, "cbam_factor": 0.900, "year": 2028, "period_code": "(2)"
    },
    "2030 (Orta Vade | Ceza: +%30 | CBAM Faktörü: %51,5)": {
        "std_markup": 1.30, "fert_markup": 1.01, "cbam_factor": 0.515, "year": 2030, "period_code": "(2)"
    },
    "2034 (Tam Uygulama | Ceza: +%30 | CBAM Faktörü: %0,0)": {
        "std_markup": 1.30, "fert_markup": 1.01, "cbam_factor": 0.000, "year": 2034, "period_code": "(2)"
    },
}

# --- TÜRKİYE SEKTÖREL TERMODİNAMİK ANOMALİ & PERSENTİL VERİ TABANI ---
TR_THERMODYNAMIC_BENCHMARKS = {
    "Demir-Çelik|(E)": {"label": "Türkiye Hurda / EAF Çelik & Hadde", "elec_p10": 0.12, "elec_p50": 0.48, "elec_p90": 0.72, "fuel_p10": 0.80, "fuel_p50": 1.65, "fuel_p90": 2.60},
    "Demir-Çelik|(H)": {"label": "Türkiye Düşük Alaşımlı EAF Çelik", "elec_p10": 0.15, "elec_p50": 0.52, "elec_p90": 0.78, "fuel_p10": 0.90, "fuel_p50": 1.80, "fuel_p90": 2.80},
    "Demir-Çelik|(C)": {"label": "Türkiye Entegre Yüksek Fırın (BF/BOF)", "elec_p10": 0.18, "elec_p50": 0.35, "elec_p90": 0.55, "fuel_p10": 14.5, "fuel_p50": 18.2, "fuel_p90": 22.0},
    "Demir-Çelik|GENEL": {"label": "Türkiye Metal İşleme / Bağlantı Elemanı / Boru", "elec_p10": 0.10, "elec_p50": 0.32, "elec_p90": 0.65, "fuel_p10": 0.40, "fuel_p50": 1.25, "fuel_p90": 2.30},
    "Alüminyum|(L)": {"label": "Türkiye İkincil Alüminyum / Ekstrüzyon Profil", "elec_p10": 0.25, "elec_p50": 0.68, "elec_p90": 1.15, "fuel_p10": 1.20, "fuel_p50": 2.40, "fuel_p90": 3.80},
    "Alüminyum|(K)": {"label": "Birincil Alüminyum Elektroliz", "elec_p10": 13.2, "elec_p50": 14.5, "elec_p90": 16.0, "fuel_p10": 1.00, "fuel_p50": 1.80, "fuel_p90": 2.80},
    "Çimento|GENEL": {"label": "Türkiye Çimento & Klinker Tesisleri", "elec_p10": 0.08, "elec_p50": 0.11, "elec_p90": 0.14, "fuel_p10": 2.90, "fuel_p50": 3.45, "fuel_p90": 4.10},
    "Gübre|GENEL": {"label": "Türkiye Azotlu Gübre / Amonyak Tesisleri", "elec_p10": 0.15, "elec_p50": 0.35, "elec_p90": 0.65, "fuel_p10": 26.0, "fuel_p50": 31.5, "fuel_p90": 37.0},
}

def evaluate_thermodynamic_percentile(sector, route_tag, elec_mwh_t, fuel_gj_t):
    key = f"{sector}|{route_tag}" if f"{sector}|{route_tag}" in TR_THERMODYNAMIC_BENCHMARKS else (
        f"{sector}|GENEL" if f"{sector}|GENEL" in TR_THERMODYNAMIC_BENCHMARKS else "Demir-Çelik|GENEL"
    )
    bm = TR_THERMODYNAMIC_BENCHMARKS[key]

    def calc_pct(val, p10, p50, p90):
        if val <= p10:
            return max(1, int(round((val / max(p10, 1e-6)) * 10)))
        elif val <= p50:
            return int(round(10 + ((val - p10) / max(p50 - p10, 1e-6)) * 40))
        elif val <= p90:
            return int(round(50 + ((val - p50) / max(p90 - p50, 1e-6)) * 40))
        else:
            over = min(9, int(round(((val - p90) / max(p90, 1e-6)) * 20)))
            return min(99, 90 + max(1, over))

    return bm, calc_pct(elec_mwh_t, bm["elec_p10"], bm["elec_p50"], bm["elec_p90"]), calc_pct(fuel_gj_t, bm["fuel_p10"], bm["fuel_p50"], bm["fuel_p90"])

# --- ZERO-KNOWLEDGE AES-256 ŞİFRELİ KASA ---
def _derive_key(passphrase: str, salt: bytes) -> bytes:
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=200_000)
    return base64.urlsafe_b64encode(kdf.derive(passphrase.encode("utf-8")))

def encrypt_vault_payload(payload_dict: dict, passphrase: str) -> bytes:
    salt = os.urandom(16)
    f = Fernet(_derive_key(passphrase, salt))
    token = f.encrypt(json.dumps(payload_dict, ensure_ascii=False).encode("utf-8"))
    envelope = {
        "format": "TR-CBAM-VAULT",
        "cipher": "AES-256-Fernet-PBKDF2",
        "created_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "salt_b64": base64.b64encode(salt).decode("ascii"),
        "token_b64": token.decode("ascii"),
    }
    return json.dumps(envelope, indent=2).encode("utf-8")

def decrypt_vault_payload(file_bytes: bytes, passphrase: str) -> dict:
    envelope = json.loads(file_bytes.decode("utf-8"))
    salt = base64.b64decode(envelope["salt_b64"])
    f = Fernet(_derive_key(passphrase, salt))
    return json.loads(f.decrypt(envelope["token_b64"].encode("ascii")).decode("utf-8"))

# --- RESMİ EXCEL / KATALOG YÜKLEYİCİ (570 CN KODU) ---
def _parse_num(val):
    if val is None:
        return None
    s = str(val).strip()
    if s in ["-", "", "N/A", "see below", "See below"]:
        return None
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return None

def _parse_excel_files(bm_path, dv_path):
    wb_bm = openpyxl.load_workbook(bm_path, data_only=True)
    ws_bm = wb_bm["Benchmarks"]
    wb_dv = openpyxl.load_workbook(dv_path, data_only=True)
    ws_tr = wb_dv["Türkiye"]
    ws_oth = wb_dv["_Other Countries and Territorie"]
    ws_a4 = wb_dv["Annex IV"]

    def read_dv(ws, is_a4=False):
        recs = {}
        for r in range(3, ws.max_row + 1):
            cn_raw = ws.cell(r, 1).value
            if cn_raw is None:
                continue
            cn_str = str(cn_raw).strip()
            if cn_str in ["Cement", "Fertilisers", "Aluminium", "Hydrogen", "Iron and steel", "Iron & Steel"]:
                continue
            cn_digits = re.sub(r"\s+", "", cn_str)
            if is_a4:
                t_em = _parse_num(ws.cell(r, 3).value)
                recs[cn_digits] = {"total": t_em, "direct": t_em, "indirect": 0.0, "route": str(ws.cell(r, 4).value or "").strip()}
            else:
                recs[cn_digits] = {
                    "direct": _parse_num(ws.cell(r, 3).value),
                    "indirect": _parse_num(ws.cell(r, 4).value),
                    "total": _parse_num(ws.cell(r, 5).value),
                    "route": str(ws.cell(r, 6).value or "").strip(),
                }
        return recs

    tr_dv = read_dv(ws_tr)
    oth_dv = read_dv(ws_oth)
    a4_dv = read_dv(ws_a4, is_a4=True)

    def resolve(cn_8, m1, m2=None):
        for p in [cn_8, cn_8[:6], cn_8[:4]]:
            for k, v in m1.items():
                if (k == p or (len(p) == 8 and k.startswith(p))) and v["total"] is not None:
                    return v, "TR"
        if m2:
            for p in [cn_8, cn_8[:6], cn_8[:4]]:
                for k, v in m2.items():
                    if (k == p or (len(p) == 8 and k.startswith(p))) and v["total"] is not None:
                        return v, "OTH"
        return None, "NONE"

    bm_rows = {}
    curr_cn, curr_desc, curr_sec = None, None, None
    sec_names = {"Cement", "Hydrogen", "Fertilisers", "Iron & Steel", "Iron and steel", "Aluminium", "Electricity"}
    for r in range(2, ws_bm.max_row + 1):
        c1 = ws_bm.cell(r, 1).value
        if c1 is not None and str(c1).strip() in sec_names:
            curr_sec = str(c1).strip()
            continue
        if c1 is not None:
            curr_cn = str(c1).strip()
            curr_desc = str(ws_bm.cell(r, 2).value or "").strip()
            bm_rows[curr_cn] = {"desc": curr_desc, "sector": curr_sec, "a": [], "b": []}
        if curr_cn:
            va, ia = ws_bm.cell(r, 3).value, ws_bm.cell(r, 4).value
            vb, ib = ws_bm.cell(r, 5).value, ws_bm.cell(r, 6).value
            if va is not None:
                bm_rows[curr_cn]["a"].append([float(va), str(ia or "").strip()])
            if vb is not None:
                bm_rows[curr_cn]["b"].append([float(vb), str(ib or "").strip()])

    sec_tr = {"Cement": "Çimento", "Hydrogen": "Hidrojen", "Fertilisers": "Gübre", "Iron & Steel": "Demir-Çelik", "Iron and steel": "Demir-Çelik", "Aluminium": "Alüminyum"}
    catalog = {}
    for cn_8, binfo in bm_rows.items():
        dv_rec, src = resolve(cn_8, tr_dv, oth_dv)
        a4_rec, _ = resolve(cn_8, a4_dv)
        catalog[cn_8] = {
            "d": binfo["desc"],
            "s": sec_tr.get(binfo["sector"], binfo["sector"]),
            "dv": dv_rec["total"] if dv_rec else 0.0,
            "dv_d": dv_rec["direct"] if dv_rec and dv_rec["direct"] is not None else (dv_rec["total"] if dv_rec else 0.0),
            "dv_i": dv_rec["indirect"] if dv_rec and dv_rec["indirect"] is not None else 0.0,
            "src": src,
            "dr": dv_rec["route"] if dv_rec else "",
            "a4": a4_rec["total"] if a4_rec else (dv_rec["total"] if dv_rec else 0.0),
            "a": binfo["a"],
            "b": binfo["b"],
        }
    return catalog

@st.cache_data
def load_full_catalog():
    if os.path.exists("CBAM Benchmarks_20260206.xlsx") and os.path.exists("DV correcting act_final update_06.08.xlsx"):
        return _parse_excel_files("CBAM Benchmarks_20260206.xlsx", "DV correcting act_final update_06.08.xlsx"), "570 CN Resmi Veri Tabanı"
    if os.path.exists("cbam_catalog_b64.txt"):
        with open("cbam_catalog_b64.txt", "r", encoding="utf-8") as f:
            raw = gzip.decompress(base64.b64decode(f.read().strip())).decode("utf-8")
        return json.loads(raw), "570 CN Resmi Veri Tabanı"
    return {
        "72085120": {"d": "Flat-rolled products of iron or non-alloy steel, hot-rolled", "s": "Demir-Çelik", "dv": 2.428, "dv_d": 2.428, "dv_i": 0.0, "src": "TR", "dr": "(C)", "a4": 8.23, "a": [[0.044, ""]], "b": [[1.37, "(C)"], [0.481, "(D)"], [0.072, "(E)"]]}
    }, "Acil Yedek Mod"

CATALOG, CATALOG_SOURCE = load_full_catalog()

def format_cn(cn_8):
    return f"{cn_8[:4]} {cn_8[4:6]} {cn_8[6:8]}" if len(cn_8) == 8 else cn_8

def extract_routes_for_item(item):
    routes = []
    for _, ind in item.get("b", []) + item.get("a", []):
        for code in re.findall(r"\(([A-L])\)", ind):
            tag = f"({code})"
            if tag not in routes:
                routes.append(tag)
    return routes if routes else ["GENEL"]

def select_benchmark_value(bm_list, route_tag, period_code):
    if not bm_list:
        return 0.0
    for val, ind in bm_list:
        if (route_tag == "GENEL" or route_tag in ind) and period_code in ind:
            return float(val)
    for val, ind in bm_list:
        if route_tag != "GENEL" and route_tag in ind:
            return float(val)
    for val, ind in bm_list:
        if period_code in ind:
            return float(val)
    return float(bm_list[0][0])

# --- AKILLI FATURA OCR ---
def parse_energy_text(text):
    kwh_matches = re.findall(r"([\d\.]+(?:,\d+)?)\s*(?:kWh|KWH|kwh)", text)
    m3_matches = re.findall(r"([\d\.]+(?:,\d+)?)\s*(?:m3|m³|Sm3|SM3)", text)
    def to_f(v):
        try:
            return float(v.replace(".", "").replace(",", "."))
        except ValueError:
            return 0.0
    kwh_vals = [to_f(x) for x in kwh_matches if to_f(x) > 10]
    m3_vals = [to_f(x) for x in m3_matches if to_f(x) > 5]
    return (max(kwh_vals) if kwh_vals else None), (max(m3_vals) if m3_vals else None)

def extract_energy_from_file(uploaded_file):
    try:
        nl = uploaded_file.name.lower()
        if nl.endswith(".pdf"):
            reader = PdfReader(uploaded_file)
            txt = "\n".join([p.extract_text() or "" for p in reader.pages])
            return parse_energy_text(txt)
        elif OCR_AVAILABLE and any(nl.endswith(e) for e in [".png", ".jpg", ".jpeg"]):
            ocr = RapidOCR()
            res, _ = ocr(np.array(Image.open(uploaded_file).convert("RGB")))
            if res:
                return parse_energy_text("\n".join([l[1] for l in res]))
        return None, None
    except Exception:
        return None, None

# --- AKREDİTE DOĞRULAYICI (VERIFIER) ÖN-KONTROL MATRİSİ ---
def run_verifier_audit_matrix(
    cn_code, sector, route_tag, eu_tonnage, total_prod_ton,
    fuels_dict, process_em_ton, elec_kwh, has_ppa,
    precursors_list, func_ratio, res_dict, meta_dict
):
    findings = []
    eff_prod = max(eu_tonnage, total_prod_ton)
    elec_mwh_per_t = (elec_kwh / 1000.0) / eff_prod
    fuel_gj_per_t = sum(qty * FUEL_FACTORS[k]["gj_per_unit"] for k, qty in fuels_dict.items()) / eff_prod
    bm_thermo, elec_pct, fuel_pct = evaluate_thermodynamic_percentile(sector, route_tag, elec_mwh_per_t, fuel_gj_per_t)

    if total_prod_ton < eu_tonnage:
        findings.append({"code": "V-01", "level": "CRITICAL", "title": "Üretim / İhracat Tonaj Çelişkisi", "detail": f"Tesis toplam üretimi ({total_prod_ton:,.1f} t), AB ihracatından ({eu_tonnage:,.1f} t) küçük olamaz."})
    else:
        findings.append({"code": "V-01", "level": "PASS", "title": "Üretim Paylaştırma Sınırı Tutarlı", "detail": f"AB sevkiyatı toplam üretimin %{res_dict['eu_share_pct']:.1f}'ini oluşturuyor."})

    if fuel_pct >= 90:
        findings.append({"code": "V-02", "level": "CRITICAL", "title": f"Termodinamik Yakıt Anomalisi (Türkiye P{fuel_pct})", "detail": f"Yakıt yoğunluğu ({fuel_gj_per_t:.2f} GJ/t), {bm_thermo['label']} tesislerinin %{fuel_pct}'inden yüksek (Medyan: {bm_thermo['fuel_p50']} GJ/t)."})
    elif elec_pct >= 90:
        findings.append({"code": "V-03", "level": "WARNING", "title": f"Yüksek Elektrik Yoğunluğu (Türkiye P{elec_pct})", "detail": f"Elektrik yoğunluğu ({elec_mwh_per_t:.2f} MWh/t), benzer tesislerin %{elec_pct}'inden yüksek (Medyan: {bm_thermo['elec_p50']} MWh/t)."})
    else:
        findings.append({"code": "V-02", "level": "PASS", "title": f"Türkiye Sektörel Termodinamik Uyum (Elektrik P{elec_pct} | Yakıt P{fuel_pct})", "detail": f"{bm_thermo['label']} referans aralığı içinde."})

    if sector == "Demir-Çelik":
        if route_tag in ["(E)", "(H)"] and fuels_dict.get("kok_komuru", 0) > 0:
            findings.append({"code": "V-07", "level": "CRITICAL", "title": "EAF Rotasında Kok Kömürü Çelişkisi", "detail": "Hurda/EAF rotasında kok kömürü girişi denetimde doğrudan ret sebebidir."})
        if route_tag in ["(C)", "(F)"] and fuels_dict.get("kok_komuru", 0) == 0 and fuels_dict.get("tas_komuru", 0) == 0:
            findings.append({"code": "V-08", "level": "CRITICAL", "title": "Yüksek Fırın (BF/BOF) Yakıt Eksikliği", "detail": "BF/BOF rotasında indirgeyici kok/taş kömürü girilmedi."})

    if sector == "Alüminyum" and route_tag == "(K)" and elec_mwh_per_t < 11.0:
        findings.append({"code": "V-09", "level": "CRITICAL", "title": "Birincil Alüminyum (K) Elektroliz Enerji Eksikliği", "detail": f"Birincil elektroliz min. 13 MWh/t gerektirir (Girilen: {elec_mwh_per_t:.2f} MWh/t)."})

    if sector == "Çimento" and process_em_ton <= 0 and not cn_code.startswith("252329"):
        findings.append({"code": "V-11", "level": "CRITICAL", "title": "Klinker Kalsinasyon Proses Emisyonu Eksik", "detail": "Çimento üretiminde CaCO₃ kalsinasyon proses emisyonu 0 girilemez."})

    is_downstream = cn_code[:4] in ["7208", "7209", "7210", "7211", "7212", "7213", "7214", "7215", "7216", "7217", "7304", "7305", "7306", "7307", "7308", "7318", "7326", "7604", "7605", "7606", "7607", "7608", "7610"]
    active_precs = [p for p in precursors_list if p["m_i"] > 0]
    if is_downstream and not active_precs:
        findings.append({"code": "V-13", "level": "CRITICAL", "title": "Karmaşık Üründe Öncül Hammadde (Precursor) Eksik", "detail": f"{format_cn(cn_code)} nihai mamuldür; öncül hammadde (m_i >= 1,00 t/t) zorunludur."})
    for p in active_precs:
        if p["m_i"] < 1.00:
            findings.append({"code": "V-14", "level": "CRITICAL", "title": "Kütlenin Korunumu İhlali (m_i < 1,00)", "detail": f"1 ton mamul için {p['m_i']:.2f} ton öncül girilmiş; m_i >= 1,00 olmalıdır."})

    if "PENDING" in meta_dict["verifier_id"].upper() or not meta_dict["verifier_id"].strip():
        findings.append({"code": "V-17", "level": "WARNING", "title": "Akredite Doğrulayıcı Ataması Bekleniyor", "detail": "Kesin beyanda akredite doğrulayıcı ID ((EU) 2025/2547) zorunludur."})
    if res_dict["actual_see_per_ton"] > res_dict["default_ef_markup"]:
        findings.append({"code": "V-19", "level": "WARNING", "title": "Tesis Emisyonu AB Default Değerini Aşıyor", "detail": f"Gerçek SEE ({res_dict['actual_see_per_ton']:.3f} t/t) AB cezalı varsayılan değerden ({res_dict['default_ef_markup']:.3f} t/t) yüksek."})

    crit_count = sum(1 for f in findings if f["level"] == "CRITICAL")
    warn_count = sum(1 for f in findings if f["level"] == "WARNING")
    pass_count = sum(1 for f in findings if f["level"] == "PASS")
    readiness_score = max(0, min(100, int(100 - (crit_count * 25) - (warn_count * 6))))
    return findings, readiness_score, crit_count, warn_count, pass_count, bm_thermo, elec_pct, fuel_pct

# --- HESAPLAMA MOTORU ---
def calculate_cbam_core(
    cn_code, route_tag, eu_tonnage, total_prod_ton,
    fuels_dict, process_em_ton, elec_kwh, grid_ef,
    precursors_list, func_ratio,
    cbam_cert_price, local_cp_eff, yr_rule
):
    item = CATALOG[cn_code]
    sector = item["s"]
    period_code = yr_rule["period_code"]
    cbam_factor = yr_rule["cbam_factor"]

    markup = yr_rule["fert_markup"] if sector == "Gübre" else yr_rule["std_markup"]
    markup_pct = int(round((markup - 1.0) * 100))
    effective_func_ratio = max(0.01, min(1.0, func_ratio))

    eff_total_prod = max(eu_tonnage, total_prod_ton)
    eu_share = eu_tonnage / eff_total_prod

    total_fuel_em = sum(qty * FUEL_FACTORS[fkey]["ef"] for fkey, qty in fuels_dict.items())
    total_facility_direct = total_fuel_em + process_em_ton
    ae_g_per_ton = total_facility_direct / eff_total_prod

    elec_mwh_per_ton = (elec_kwh / 1000.0) / eff_total_prod
    fuel_gj_per_ton = sum(qty * FUEL_FACTORS[fkey]["gj_per_unit"] for fkey, qty in fuels_dict.items()) / eff_total_prod
    indirect_per_ton = elec_mwh_per_ton * grid_ef
    is_indirect_billable = sector in ["Çimento", "Gübre"]

    prec_see_sum_per_ton = 0.0
    prec_sefa_sum_per_ton = 0.0
    prec_details = []

    for p in precursors_list:
        m_i = float(p["m_i"])
        if m_i <= 0:
            continue
        p_cn = p["cn"]
        p_origin = p["origin"]
        p_verified = p["verified"]
        p_actual_see = float(p["actual_see"])
        p_route = p["route"]

        if p_origin in ["AB / EFTA Menşeli", "Hurda / Sıfır Gömülü Emisyon"]:
            see_i = 0.0
            sefa_i = 0.0
        else:
            p_item = CATALOG[p_cn]
            if p_verified:
                see_i = p_actual_see
                bm_p = select_benchmark_value(p_item["b"], p_route, period_code)
                sefa_i = cbam_factor * 1.0 * bm_p
            else:
                base_p_dv = p_item["a4"] if "Bilinmeyen" in p_origin else p_item["dv"]
                p_markup = yr_rule["fert_markup"] if p_item["s"] == "Gübre" else yr_rule["std_markup"]
                see_i = base_p_dv * p_markup
                def_p_route = p_item["dr"] if p_item["dr"] else p_route
                bm_p = select_benchmark_value(p_item["b"], def_p_route, period_code)
                sefa_i = cbam_factor * 1.0 * bm_p

        prec_see_sum_per_ton += m_i * see_i
        prec_sefa_sum_per_ton += m_i * sefa_i
        prec_details.append({"cn": format_cn(p_cn), "m_i": m_i, "origin": p_origin, "see_i": see_i, "sefa_i": sefa_i})

    actual_see_per_ton = (ae_g_per_ton + (indirect_per_ton if is_indirect_billable else 0.0) + prec_see_sum_per_ton) * effective_func_ratio
    actual_total_em = actual_see_per_ton * eu_tonnage

    col_a_bm = select_benchmark_value(item["a"], route_tag, period_code)
    sfa_proc_per_ton = cbam_factor * 1.0 * col_a_bm
    sefa_actual_per_ton = (sfa_proc_per_ton + prec_sefa_sum_per_ton) * effective_func_ratio
    faa_actual_total = sefa_actual_per_ton * eu_tonnage

    default_ef_with_markup = item["dv"] * markup * effective_func_ratio
    default_total_em = default_ef_with_markup * eu_tonnage

    default_route = item["dr"] if item["dr"] else route_tag
    col_b_bm = select_benchmark_value(item["b"], default_route, period_code)
    sefa_default_per_ton = cbam_factor * 1.0 * col_b_bm * effective_func_ratio
    faa_default_total = sefa_default_per_ton * eu_tonnage

    cert_default = max(0.0, default_total_em - faa_default_total)
    cert_actual = max(0.0, actual_total_em - faa_actual_total)

    cp_reduction_default = min(cert_default, (local_cp_eff / cbam_cert_price) * cert_default) if cbam_cert_price > 0 else 0.0
    cp_reduction_actual = min(cert_actual, (local_cp_eff / cbam_cert_price) * cert_actual) if cbam_cert_price > 0 else 0.0

    net_cert_default = max(0.0, cert_default - cp_reduction_default)
    net_cert_actual = max(0.0, cert_actual - cp_reduction_actual)

    tax_default = net_cert_default * cbam_cert_price
    tax_actual = net_cert_actual * cbam_cert_price
    direct_savings = max(0.0, tax_default - tax_actual)
    savings_per_ton_eur = direct_savings / eu_tonnage if eu_tonnage > 0 else 0.0

    return {
        "cn": format_cn(cn_code), "cn_raw": cn_code, "desc": item["d"], "sector": sector,
        "dv_source": "Türkiye ((EU) 2025/2621 & 2026/1740)" if item["src"] == "TR" else "Diğer Ülkeler Fallback ((EU) 2025/2621)",
        "route_tag": route_tag, "default_route": default_route,
        "markup_pct": markup_pct, "cbam_factor": cbam_factor, "func_ratio": effective_func_ratio,
        "eff_total_prod": eff_total_prod, "eu_share_pct": eu_share * 100.0,
        "elec_mwh_per_ton": elec_mwh_per_ton, "fuel_gj_per_ton": fuel_gj_per_ton,
        "ae_g_per_ton": ae_g_per_ton, "indirect_per_ton": indirect_per_ton, "is_indirect_billable": is_indirect_billable,
        "prec_see_per_ton": prec_see_sum_per_ton, "prec_sefa_per_ton": prec_sefa_sum_per_ton, "prec_details": prec_details,
        "actual_see_per_ton": actual_see_per_ton, "actual_total_em": actual_total_em,
        "default_ef_markup": default_ef_with_markup, "default_total_em": default_total_em,
        "col_a_bm": col_a_bm, "col_b_bm": col_b_bm,
        "faa_actual_total": faa_actual_total, "faa_default_total": faa_default_total,
        "cert_default": cert_default, "cert_actual": cert_actual,
        "cp_reduction_actual": cp_reduction_actual,
        "net_cert_default": net_cert_default, "net_cert_actual": net_cert_actual,
        "tax_default": tax_default, "tax_actual": tax_actual, "direct_savings": direct_savings,
        "savings_per_ton_eur": savings_per_ton_eur,
        "is_de_minimis_candidate": (eu_tonnage < 50.0) and (cn_code != "28041000"),
    }

# --- AB KOMİSYONU RESMİ COMMUNICATION TEMPLATE EXCEL ÜRETİCİSİ ---
def generate_eu_communication_template_xlsx(meta, res):
    wb = openpyxl.Workbook()
    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    sub_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    accent_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    white_bold = Font(color="FFFFFF", bold=True, size=11)
    bold_font = Font(bold=True, size=10)
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1")
    )

    ws1 = wb.active
    ws1.title = "Summary_Communication"
    ws1.merge_cells("A1:D1")
    ws1["A1"] = f"EU CBAM INSTALLATION OPERATOR COMMUNICATION TEMPLATE | {meta['broker_name'].upper()}"
    ws1["A1"].fill = header_fill
    ws1["A1"].font = white_bold
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 28

    rows_s1 = [
        ("1. INSTALLATION IDENTIFICATION", "", "", ""),
        ("Installation Operator Name", meta["company_name"], "CBAM Registry Installation ID", meta["installation_id"]),
        ("Country of Installation", "Türkiye (TR)", "UN/LOCODE", meta["unlocode"]),
        ("Geographic Coordinates", meta["coordinates"], "Reporting Period (Year)", str(meta["year"])),
        ("Accredited Verifier ID", meta["verifier_id"], "Advisory / Representative", meta["broker_name"]),
        ("", "", "", ""),
        ("2. AGGREGATED GOODS & SPECIFIC EMBEDDED EMISSIONS (SEE)", "", "", ""),
        ("CN Code", res["cn"], "Product Description", res["desc"][:80]),
        ("Sector", res["sector"], "Production Route Indicator", res["route_tag"]),
        ("Total Facility Production (t)", round(res["eff_total_prod"], 3), "EU Export Mass (t)", round(meta["tonnage"], 3)),
        ("Direct Process SEE (ae_g) [tCO2e/t]", round(res["ae_g_per_ton"], 4), "Indirect Electricity SEE [tCO2e/t]", round(res["indirect_per_ton"], 4)),
        ("Precursor Embedded SEE [tCO2e/t]", round(res["prec_see_per_ton"], 4), "Total Specific Embedded Emissions (SEE_g) [tCO2e/t]", round(res["actual_see_per_ton"], 4)),
        ("", "", "", ""),
        ("3. ARTICLE 31 FREE ALLOCATION & CERTIFICATE SUMMARY", "", "", ""),
        ("Column A Process Benchmark (BM_g*) [tCO2e/t]", round(res["col_a_bm"], 3), "CBAM Factor (Article 31)", res["cbam_factor"]),
        ("Total Free Allocation Adjustment (FAA) [tCO2e]", round(res["faa_actual_total"], 4), "Net CBAM Certificates to Surrender", round(res["net_cert_actual"], 4)),
        ("Estimated Net Liability (EUR)", round(res["tax_actual"], 2), "Avoided Default Penalty Savings (EUR)", round(res["direct_savings"], 2)),
    ]

    for idx, rdata in enumerate(rows_s1, start=3):
        for c_i, val in enumerate(rdata, start=1):
            cell = ws1.cell(idx, c_i, val)
            cell.border = thin_border
            if rdata[1] == "" and rdata[0] != "":
                cell.fill = sub_fill
                cell.font = bold_font
            elif c_i in [1, 3]:
                cell.font = bold_font
            if "Savings" in str(val) or (idx == 19 and c_i == 4):
                cell.fill = accent_fill

    for col_letter, width in [("A", 42), ("B", 32), ("C", 42), ("D", 32)]:
        ws1.column_dimensions[col_letter].width = width

    ws2 = wb.create_sheet("Precursors_&_Routes")
    headers_s2 = ["Precursor CN Code", "Specific Mass m_i (t/t)", "Origin / Status", "Precursor SEE_i (tCO2e/t)", "Precursor SEFA_i (tCO2e/t)"]
    for c_i, h in enumerate(headers_s2, start=1):
        c = ws2.cell(1, c_i, h)
        c.fill = header_fill
        c.font = white_bold
        c.border = thin_border
        ws2.column_dimensions[openpyxl.utils.get_column_letter(c_i)].width = 28

    for r_i, p in enumerate(res.get("prec_details", []), start=2):
        vals = [p["cn"], p["m_i"], p["origin"], round(p["see_i"], 4), round(p["sefa_i"], 4)]
        for c_i, v in enumerate(vals, start=1):
            cell = ws2.cell(r_i, c_i, v)
            cell.border = thin_border

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# --- AB ALICISINA ÖZEL İNGİLİZCE SİPARİŞ & MARJ KORUMA TEKLİF PDF'İ ---
def setup_unicode_font(pdf):
    if os.path.exists("DejaVuSans.ttf"):
        pdf.add_font("TrFont", "", "DejaVuSans.ttf")
        pdf.add_font("TrFont", "B", "DejaVuSans-Bold.ttf" if os.path.exists("DejaVuSans-Bold.ttf") else "DejaVuSans.ttf")
        pdf.add_font("TrFont", "I", "DejaVuSans-Oblique.ttf" if os.path.exists("DejaVuSans-Oblique.ttf") else "DejaVuSans.ttf")
        return "TrFont"
    return "Helvetica"

def generate_eu_buyer_pitch_pdf(meta, res, buyer_company="EU Importer / Customer GmbH"):
    pdf = FPDF()
    pdf.add_page()
    font_name = setup_unicode_font(pdf)

    pdf.set_fill_color(15, 23, 42)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(font_name, "B", 11)
    pdf.cell(0, 11, "EU CBAM VERIFIED COST ADVANTAGE & SUPPLIER COMPLIANCE BRIEF", ln=True, align="C", fill=True)

    pdf.set_text_color(71, 85, 105)
    pdf.set_font(font_name, "", 8.2)
    pdf.cell(0, 6, f"Prepared by: {meta['company_name']} (Türkiye)  |  For: {buyer_company}  |  Reporting Year: {meta['year']}", ln=True, align="C")
    pdf.ln(4)

    pdf.set_text_color(15, 23, 42)
    pdf.set_font(font_name, "B", 10)
    pdf.cell(0, 6, "EXECUTIVE SUMMARY FOR EU IMPORT PROCUREMENT & COMPLIANCE", ln=True)
    pdf.set_font(font_name, "", 8.8)
    pdf.multi_cell(0, 5, f"Under the EU Carbon Border Adjustment Mechanism (CBAM) Definitive Regime (Reg. (EU) 2023/956 & IR (EU) 2026/1740), EU importers declaring goods without installation-specific emission data face punitive default values (+{res['markup_pct']}% mark-up in {meta['year']}). By sourcing from {meta['company_name']}, your import customs declaration benefits from installation-specific primary emission tracking:")
    pdf.ln(3)

    pdf.set_fill_color(241, 245, 249)
    pdf.set_font(font_name, "B", 8.5)
    pdf.cell(85, 7, "CBAM Declaration Basis", border=1, fill=True)
    pdf.cell(35, 7, "SEE (tCO2e/t)", border=1, align="C", fill=True)
    pdf.cell(35, 7, "Net Certificates", border=1, align="C", fill=True)
    pdf.cell(35, 7, "CBAM Cost (EUR)", border=1, align="R", fill=True, ln=True)

    pdf.set_font(font_name, "", 8.5)
    pdf.cell(85, 7, f"EU Default Value (+{res['markup_pct']}% Mark-up)", border=1)
    pdf.cell(35, 7, f"{res['default_ef_markup']:.3f}", border=1, align="C")
    pdf.cell(35, 7, f"{res['net_cert_default']:,.2f}", border=1, align="C")
    pdf.cell(35, 7, f"{res['tax_default']:,.2f} EUR", border=1, align="R", ln=True)

    pdf.cell(85, 7, f"{meta['company_name']} Actual Data", border=1)
    pdf.cell(35, 7, f"{res['actual_see_per_ton']:.3f}", border=1, align="C")
    pdf.cell(35, 7, f"{res['net_cert_actual']:,.2f}", border=1, align="C")
    pdf.cell(35, 7, f"{res['tax_actual']:,.2f} EUR", border=1, align="R", ln=True)

    pdf.set_font(font_name, "B", 9.5)
    pdf.set_fill_color(220, 252, 231)
    pdf.cell(155, 8.5, f"DIRECT CBAM IMPORT COST SAVING FOR {buyer_company.upper()} ({meta['tonnage']:,.1f} t)", border=1, fill=True)
    pdf.cell(35, 8.5, f"+{res['direct_savings']:,.2f} EUR", border=1, align="R", fill=True, ln=True)
    pdf.ln(3)

    pdf.set_font(font_name, "B", 9)
    pdf.cell(0, 6, f"Unit Import Cost Advantage: +{res['savings_per_ton_eur']:,.2f} EUR saved per tonne imported", ln=True)
    pdf.set_font(font_name, "", 8.2)
    pdf.cell(0, 5, f"• CN Code: {res['cn']} ({res['desc'][:65]})", ln=True)
    pdf.cell(0, 5, f"• Installation ID: {meta['installation_id']} | UN/LOCODE: {meta['unlocode']} | Production Route: {res['route_tag']}", ln=True)
    pdf.cell(0, 5, f"• Advisory & Verification Partner: {meta['broker_name']}", ln=True)

    return bytes(pdf.output())

# --- O3CI v3.2 XML ÜRETİCİ ---
def generate_cbam_xml_core(meta, results_list):
    root = ET.Element("CBAMDefinitiveDeclaration", xmlns="urn:eu:taxud:cbam:o3ci:v3.2")
    header = ET.SubElement(root, "DeclarationHeader")
    ET.SubElement(header, "ReportingYear").text = str(meta["year"])
    ET.SubElement(header, "GeneratedAtUTC").text = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    ET.SubElement(header, "AuthorisedDeclarantRepresentative").text = meta["broker_name"]
    ET.SubElement(header, "AccreditedVerifierID").text = meta["verifier_id"]
    ET.SubElement(header, "LegalBasis").text = "Reg (EU) 2023/956; Reg (EU) 2025/2083; IR (EU) 2025/2547; IR (EU) 2025/2620; IR (EU) 2025/2621; IR (EU) 2026/1740; IR (EU) 2026/2498"

    inst = ET.SubElement(root, "InstallationDetails")
    ET.SubElement(inst, "InstallationOperatorName").text = meta["company_name"]
    ET.SubElement(inst, "CBAMRegistryInstallationID").text = meta["installation_id"]
    ET.SubElement(inst, "CountryCode").text = "TR"
    ET.SubElement(inst, "UNLOCODE").text = meta["unlocode"]
    ET.SubElement(inst, "GeographicCoordinates").text = meta["coordinates"]

    goods_el = ET.SubElement(root, "CBAMImportedGoods")
    for idx, item in enumerate(results_list, start=1):
        good = ET.SubElement(goods_el, "GoodItem", sequence=str(idx))
        ET.SubElement(good, "CNCode").text = item["cn_raw"]
        ET.SubElement(good, "Sector").text = item["sector"]
        ET.SubElement(good, "ProductionRouteCode").text = item["route_tag"]
        ET.SubElement(good, "NetMassTonnes").text = f"{item['tonnage']:.3f}"
        ET.SubElement(good, "FunctionalUnitRatio").text = f"{item['func_ratio']:.4f}"
        ET.SubElement(good, "DirectProcessSEE_tCO2e_per_t").text = f"{item['ae_g_per_ton']:.4f}"
        ET.SubElement(good, "PrecursorSEE_tCO2e_per_t").text = f"{item['prec_see_per_ton']:.4f}"
        ET.SubElement(good, "TotalActualSEE_tCO2e_per_t").text = f"{item['actual_see_per_ton']:.4f}"
        ET.SubElement(good, "TotalActualEmbeddedEmissions_tCO2e").text = f"{item['actual_total_em']:.4f}"
        ET.SubElement(good, "ColumnAProcessBenchmark_BMg_star").text = f"{item['col_a_bm']:.3f}"
        ET.SubElement(good, "FreeAllocationAdjustment_FAA_tCO2e").text = f"{item['faa_actual_total']:.4f}"
        ET.SubElement(good, "ThirdCountryCarbonPriceDeduction_Certs").text = f"{item['cp_reduction_actual']:.4f}"
        ET.SubElement(good, "NetCBAMCertificatesToSurrender").text = f"{item['net_cert_actual']:.4f}"
        ET.SubElement(good, "EstimatedCertificateLiability_EUR").text = f"{item['tax_actual']:.2f}"
        ET.SubElement(good, "AvoidedDefaultPenalty_EUR").text = f"{item['direct_savings']:.2f}"

    reparsed = minidom.parseString(ET.tostring(root, "utf-8"))
    return reparsed.toprettyxml(indent="  ").encode("utf-8")

# --- WHITE-LABEL KURUMSAL PRE-AUDIT PDF ÜRETİCİ ---
def generate_corporate_pdf_core(meta, res, audit_summary, is_licensed=False, logo_bytes=None):
    pdf = FPDF()
    pdf.add_page()
    font_name = setup_unicode_font(pdf)

    if not is_licensed:
        pdf.set_fill_color(254, 226, 226)
        pdf.set_text_color(185, 28, 28)
        pdf.set_font(font_name, "B", 8.2)
        pdf.cell(0, 7, "TASLAK ÖN-ANALİZ ÇIKTISI — RESMİ BEYAN VE ALICI TEKLİFİ İÇİN GEÇERSİZDİR (LİSANS GEREKLİDİR)", ln=True, align="C", fill=True)
        pdf.ln(2)

    temp_logo = None
    if logo_bytes:
        try:
            img = Image.open(io.BytesIO(logo_bytes))
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                img.save(tmp.name, format="PNG")
                temp_logo = tmp.name
            pdf.image(temp_logo, x=12, y=18 if not is_licensed else 10, w=24)
            pdf.set_xy(38, 20 if not is_licensed else 12)
        except Exception:
            pass

    pdf.set_fill_color(15, 23, 42)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(font_name, "B", 10.6)
    pdf.cell(0, 10.5, f"{meta['broker_name'].upper()} | AB SKDM KESİN DÖNEM DENETİM RAPORU", ln=True, align="C", fill=True)

    pdf.set_text_color(80, 80, 80)
    pdf.set_font(font_name, "", 7.8)
    pdf.cell(0, 5.5, f"Tarih: {datetime.now().strftime('%d.%m.%Y %H:%M')}  |  Tesis ID: {meta['installation_id']}  |  Denetim Hazırlık Skoru: %{audit_summary['score']}", ln=True, align="C")
    pdf.ln(3)

    pdf.set_text_color(0, 0, 0)
    pdf.set_font(font_name, "B", 9.3)
    pdf.cell(0, 5.8, "1. TESİS KİMLİĞİ, ÜRETİM ROTASI VE KAPSAM BİLGİLERİ", ln=True)
    pdf.set_font(font_name, "", 8.2)
    pdf.cell(0, 4.7, f"İhracatçı Kuruluş: {meta['company_name']}  |  UN/LOCODE: {meta['unlocode']}  |  Koordinat: {meta['coordinates']}", ln=True)
    pdf.cell(0, 4.7, f"CN Kodu: {res['cn']} — {res['desc'][:66]}", ln=True)
    pdf.cell(0, 4.7, f"Sektör: {res['sector']}  |  Üretim Rotası: {ROUTE_LABELS.get(res['route_tag'], res['route_tag'])}", ln=True)
    pdf.cell(0, 4.7, f"Tesis Üretimi: {res['eff_total_prod']:,.2f} t  |  AB Sevkiyatı: {meta['tonnage']:,.2f} t (Pay: %{res['eu_share_pct']:.1f})  |  Fonk. Oran: %{res['func_ratio']*100:.1f}", ln=True)
    pdf.ln(2.5)

    pdf.set_fill_color(241, 245, 249)
    pdf.set_font(font_name, "B", 9.3)
    pdf.cell(0, 6.8, "2. AB SKDM SERTİFİKA YÜKÜMLÜLÜĞÜ KIYASLAMASI ((EU) 2025/2620 & 2026/2498)", ln=True, fill=True)
    pdf.ln(1.2)

    pdf.set_font(font_name, "B", 7.8)
    pdf.cell(72, 6.4, "Beyan Senaryosu", border=1)
    pdf.cell(38, 6.4, "Gömülü Emisyon", border=1, align="C")
    pdf.cell(40, 6.4, "Net Sertifika (Adet)", border=1, align="C")
    pdf.cell(40, 6.4, "Ödenecek Tutar", border=1, align="R", ln=True)

    pdf.set_font(font_name, "", 7.8)
    pdf.cell(72, 6.4, f"A: Kanıtsız Default (+%{res['markup_pct']} Marj | Col B)", border=1)
    pdf.cell(38, 6.4, f"{res['default_total_em']:,.2f} tCO2e", border=1, align="C")
    pdf.cell(40, 6.4, f"{res['net_cert_default']:,.2f}", border=1, align="C")
    pdf.cell(40, 6.4, f"{res['tax_default']:,.2f} EUR", border=1, align="R", ln=True)

    pdf.cell(72, 6.4, "B: Kanıtlı Gerçek Veri (Col A + Precursor)", border=1)
    pdf.cell(38, 6.4, f"{res['actual_total_em']:,.2f} tCO2e", border=1, align="C")
    pdf.cell(40, 6.4, f"{res['net_cert_actual']:,.2f}", border=1, align="C")
    pdf.cell(40, 6.4, f"{res['tax_actual']:,.2f} EUR", border=1, align="R", ln=True)

    pdf.set_font(font_name, "B", 8.9)
    pdf.set_fill_color(220, 252, 231)
    pdf.cell(150, 7.8, "DOĞRUDAN KORUNAN NET SKDM SERTİFİKA TASARRUFU (A - B)", border=1, fill=True)
    pdf.cell(40, 7.8, f"+{res['direct_savings']:,.2f} EUR", border=1, align="R", ln=True, fill=True)
    pdf.ln(2.5)

    pdf.set_font(font_name, "B", 9.3)
    pdf.cell(0, 5.8, "3. ANNEX III EMİSYON KIRILIMI & TERMODİNAMİK KIYAS ÖZETİ", ln=True)
    pdf.set_font(font_name, "", 8.1)
    pdf.cell(0, 4.6, f"• Tesis Doğrudan Emisyonu (ae_g): {res['ae_g_per_ton']:.4f} tCO2e/t  |  Öncül Hammadde (∑ m_i x SEE_i): {res['prec_see_per_ton']:.4f} tCO2e/t", ln=True)
    pdf.cell(0, 4.6, f"• Türkiye Sektör Persentil Konumu: Elektrik P{audit_summary['elec_pct']} ({res['elec_mwh_per_ton']:.3f} MWh/t) | Yakıt P{audit_summary['fuel_pct']} ({res['fuel_gj_per_ton']:.2f} GJ/t)", ln=True)
    pdf.cell(0, 4.6, f"• Gerçek Veri Ücretsiz Tahsisat (Col A BM={res['col_a_bm']:.3f} + Precursor SEFA): -{res['faa_actual_total']:,.2f} tCO2e", ln=True)
    pdf.cell(0, 4.6, f"• Verifier Ön-Kontrol Matrisi ((EU) 2025/2547): Skor %{audit_summary['score']} ({audit_summary['crit']} Kritik, {audit_summary['warn']} Uyarı)", ln=True)
    pdf.ln(2)

    pdf.set_font(font_name, "B", 8.2)
    pdf.cell(0, 5, f"Düzenleyen Kuruluş / Danışman: {meta['broker_name']}  |  Doğrulayıcı ID: {meta['verifier_id']}", ln=True)
    pdf.ln(1)
    pdf.set_font(font_name, "", 6.9)
    pdf.set_text_color(75, 75, 75)
    pdf.multi_cell(0, 3.6, "Mevzuat & Veri Güvenliği Beyanı: Bu rapor (EU) 2023/956, (EU) 2025/2083, (EU) 2025/2547, (EU) 2025/2620, (EU) 2025/2621, (EU) 2026/1740 ve (EU) 2026/2498 sayılı AB SKDM tüzüklerine uygun hazırlanmıştır. Tesis verileri istemci parolasıyla AES-256 şifreli .cbamvault dosyasında saklanır.")

    out_bytes = bytes(pdf.output())
    if temp_logo and os.path.exists(temp_logo):
        try:
            os.remove(temp_logo)
        except Exception:
            pass
    return out_bytes

# --- SESSION STATE BAŞLATMA ---
if "vault_data" not in st.session_state:
    st.session_state["vault_data"] = {
        "company_name": "Örnek Metal ve Çelik Sanayi A.Ş.",
        "buyer_company": "Muster Stahlhandel GmbH (DE)",
        "installation_id": "TR-INST-2026-0001",
        "unlocode": "TRIST",
        "coordinates": "41.0082, 28.9784",
        "selected_cn": "72085120",
        "selected_route": "(E)",
        "total_prod_ton": 1000.0,
        "eu_tonnage": 120.0,
        "q_gas": 24000.0,
        "elec_kwh": 180000.0,
        "prec_cn": "72071114",
        "prec_mi": 1.05,
        "prec_verified": False,
        "prec_actual_see": 0.45,
        "history": []
    }
if "last_audit_findings" not in st.session_state:
    st.session_state["last_audit_findings"] = []

vd = st.session_state["vault_data"]

# --- SAP FIORI SHELL BAR & SAĞ ÜST 4'LÜ KONTROL ALANI ---
shell_left, shell_ctrl1, shell_ctrl2, shell_ctrl3, shell_ctrl4 = st.columns([2.2, 1.15, 0.95, 0.65, 0.95])

with shell_ctrl1:
    selected_theme_name = st.selectbox("Tema", list(FIORI_THEMES.keys()), index=0, label_visibility="collapsed")
with shell_ctrl2:
    selected_font_name = st.selectbox("Yazı Tipi", list(FIORI_FONTS.keys()), index=0, label_visibility="collapsed")

inject_fiori_design_system(selected_theme_name, selected_font_name)
active_theme = FIORI_THEMES[selected_theme_name]

with shell_ctrl3:
    f_list = st.session_state.get("last_audit_findings", [])
    crit_n = sum(1 for x in f_list if x["level"] == "CRITICAL")
    bell_label = f"🔔 ({crit_n})" if crit_n > 0 else "🔔 (0)"
    with st.popover(bell_label, use_container_width=True):
        st.markdown("**Denetçi Ön-Kontrol Bildirimleri**")
        if not f_list:
            st.caption("Henüz analiz çalıştırılmadı. Sevkiyat analizi sonrası termodinamik ve mevzuat uyarıları burada görünür.")
        else:
            for item_f in f_list:
                if item_f["level"] == "CRITICAL":
                    st.error(f"**[{item_f['code']}] {item_f['title']}:** {item_f['detail']}")
                elif item_f["level"] == "WARNING":
                    st.warning(f"**[{item_f['code']}] {item_f['title']}:** {item_f['detail']}")
                else:
                    st.success(f"**[{item_f['code']}] {item_f['title']}:** {item_f['detail']}")

with shell_ctrl4:
    with st.popover("🏢 Profil & Lisans", use_container_width=True):
        st.markdown("**Kurumsal Lisans & Piyasa**")
        license_key = st.text_input("Lisans Anahtarı", type="password", value="", placeholder="TRCBAM-PRO-XXXX")
        is_licensed, license_label = check_license_active(license_key)
        if is_licensed:
            st.success(f"✓ {license_label}")
        else:
            st.warning(f"🔒 {license_label}")

        year_choice = st.selectbox("Beyan Yılı", list(YEAR_RULES.keys()))
        yr_rule = YEAR_RULES[year_choice]
        cbam_cert_price = st.number_input("AB ETS Karbon Fiyatı (€ / tCO₂e)", min_value=10.0, value=82.32, step=0.5)
        if yr_rule["year"] in [2026, 2027]:
            local_cp_eff = 0.0
            st.caption("🇹🇷 TR-ETS 2026–2027 Mahsup: **0,00 €/tCO₂e**")
        else:
            local_cp_eff = st.number_input("TR-ETS Mahsup (€ / tCO₂e)", min_value=0.0, value=15.0, step=1.0)

with shell_left:
    st.markdown(f"""
    <div class="fiori-shell-bar">
        <div class="fiori-brand">
            <div class="fiori-logo-box">TR-CBAM</div>
            <div>
                <div class="fiori-title">AB Sınırda Karbon (SKDM) Ceza Kalkanı</div>
                <div class="fiori-subtitle">İhracat Uyum, Termodinamik Denetim & Alıcı Teklif Platformu • ETS: <strong>{cbam_cert_price:.2f} €/tCO₂e</strong></div>
            </div>
        </div>
        <div class="fiori-compliance-badge">● AB SKDM Mevzuat Uyumlu</div>
    </div>
    """, unsafe_allow_html=True)

# --- 6 ANA MODÜL YATAY SEKME NAVİGASYONU ---
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "1. Sevkiyat & Ceza Analizi",
    "2. Gümrük Müşaviri Referral",
    "3. ERP Veri Aktarımı",
    "4. Tedarikçi (Precursor) Portalı",
    "5. Çoklu Sevkiyat & 50t Eşiği",
    "6. Resmi Katalog"
])

# ==============================================================================
# SEKME 1: SEVKİYAT & CEZA ANALİZİ (SAP FIORI MODÜLER KART IZGARASI)
# ==============================================================================
with tab1:
    col_left, col_right = st.columns([1.08, 1.82], gap="large")

    with col_left:
        st.markdown('<div class="fiori-tile">', unsafe_allow_html=True)
        ui_mode = st.radio(
            "Çalışma Modu",
            ["⚡ Hızlı Analiz Modu (4 Temel Veri)", "🔬 Detaylı Mühendislik Modu"],
            horizontal=True
        )
        is_quick_mode = ui_mode.startswith("⚡")

        search_q = st.text_input("🔍 GTİP / CN Kodu veya Ürün Adı Ara (Örn: 7208, 7318, 7604, 2523, 3102)", value="")
        filtered_cns = [
            cn for cn, info in CATALOG.items()
            if not search_q.strip() or search_q.lower() in cn.lower() or search_q.lower() in format_cn(cn).lower() or search_q.lower() in info["d"].lower()
        ] or list(CATALOG.keys())

        default_idx = filtered_cns.index(vd["selected_cn"]) if vd["selected_cn"] in filtered_cns else 0
        selected_cn = st.selectbox(
            f"1. İhracat Ürünü CN / GTİP Kodu ({len(filtered_cns)} kayıt)",
            filtered_cns,
            index=default_idx,
            format_func=lambda x: f"{format_cn(x)} | [{CATALOG[x]['s']}] {CATALOG[x]['d'][:62]}"
        )
        sel_item = CATALOG[selected_cn]

        avail_routes = extract_routes_for_item(sel_item)
        default_route = "(E)" if sel_item["s"] == "Demir-Çelik" and "(E)" in avail_routes else (
            "(L)" if sel_item["s"] == "Alüminyum" and "(L)" in avail_routes else avail_routes[0]
        )

        if is_quick_mode:
            selected_route = default_route
            func_ratio = 0.75 if sel_item["s"] == "Çimento" and not selected_cn.startswith("252310") else (0.27 if sel_item["s"] == "Gübre" else 1.0)
        else:
            default_route_idx = avail_routes.index(vd["selected_route"]) if vd["selected_route"] in avail_routes else avail_routes.index(default_route)
            selected_route = st.selectbox("Üretim Rotası ((EU) 2025/2620 Ek Bölüm 5.3)", avail_routes, index=default_route_idx, format_func=lambda r: ROUTE_LABELS.get(r, r))
            func_ratio = 1.0
            if sel_item["s"] == "Çimento" and not selected_cn.startswith("252310"):
                func_ratio = st.slider("Çimento Klinker Oranı (%)", 10, 100, 75) / 100.0
            elif sel_item["s"] == "Gübre":
                func_ratio = st.slider("Gübre Azot (N) İçeriği (%)", 5, 82, 27) / 100.0

        invoice_file = st.file_uploader("Opsiyonel: Elektrik / Doğalgaz Faturası Yükle", type=["pdf", "png", "jpg", "jpeg"])
        auto_kwh, auto_gas = float(vd["elec_kwh"]), float(vd["q_gas"])
        invoice_status = "Kanıt Belgesi Yüklenmedi"
        if invoice_file:
            dk, dg = extract_energy_from_file(invoice_file)
            invoice_status = f"Doğrulandı: {invoice_file.name}"
            if dk:
                auto_kwh = dk
                st.success(f"Elektrik okundu: **{dk:,.2f} kWh**")
            if dg:
                auto_gas = dg
                st.success(f"Doğalgaz okundu: **{dg:,.2f} m³**")

        with st.form("fiori_main_form"):
            company_name = st.text_input("İhracatçı Firma Unvanı (TR)", value=vd["company_name"])
            buyer_company = st.text_input("AB İthalatçı / Müşteri Unvanı", value=vd.get("buyer_company", "Muster Stahlhandel GmbH (DE)"))

            if is_quick_mode:
                eu_tonnage = st.number_input("2. AB Sevkiyat Net Ağırlığı (Ton)", min_value=0.1, value=float(vd["eu_tonnage"]), step=10.0)
                q_e1, q_e2 = st.columns(2)
                elec_kwh = q_e1.number_input("3. Elektrik (kWh)", min_value=0.0, value=float(auto_kwh), step=1000.0)
                q_gas = q_e2.number_input("4. Doğalgaz (m³)", min_value=0.0, value=float(auto_gas), step=500.0)

                total_prod_ton = max(eu_tonnage, float(vd["total_prod_ton"]))
                q_coal = q_coke = q_lignite = q_fueloil = q_lpg = process_em_ton = 0.0
                has_ppa = False
                active_grid_ef = EU_DEFAULT_TR_GRID_EF
                prec_cn = "72071114" if sel_item["s"] == "Demir-Çelik" else ("76011090" if sel_item["s"] == "Alüminyum" else selected_cn)
                prec_mi = 1.05 if sel_item["s"] in ["Demir-Çelik", "Alüminyum"] else 0.0
                prec_origin = "Türkiye / Diğer Üçüncü Ülke"
                prec_verified = False
                prec_actual_see = 0.45
                prec_route = selected_route
            else:
                p1, p2 = st.columns(2)
                total_prod_ton = p1.number_input("Tesis Toplam Üretimi (Ton)", min_value=1.0, value=float(vd["total_prod_ton"]), step=50.0)
                eu_tonnage = p2.number_input("AB Sevkiyatı (Ton)", min_value=0.1, value=float(vd["eu_tonnage"]), step=10.0)

                f1, f2 = st.columns(2)
                q_gas = f1.number_input("Doğalgaz (m³)", min_value=0.0, value=float(auto_gas), step=500.0)
                q_coal = f2.number_input("Taş Kömürü (Ton)", min_value=0.0, value=0.0, step=1.0)
                f3, f4 = st.columns(2)
                q_coke = f3.number_input("Kok Kömürü (Ton)", min_value=0.0, value=0.0, step=1.0)
                q_lignite = f4.number_input("Linyit (Ton)", min_value=0.0, value=0.0, step=1.0)
                f5, f6 = st.columns(2)
                q_fueloil = f5.number_input("Fuel-Oil (Ton)", min_value=0.0, value=0.0, step=1.0)
                q_lpg = f6.number_input("LPG (Ton)", min_value=0.0, value=0.0, step=1.0)

                process_em_ton = st.number_input("Proses Emisyonları (tCO₂e)", min_value=0.0, value=0.0, step=5.0)
                elec_kwh = st.number_input("Elektrik Tüketimi (kWh)", min_value=0.0, value=float(auto_kwh), step=1000.0)
                has_ppa = st.checkbox("Fiziksel PPA / Doğrudan Hat Kanıtlı", value=False)
                active_grid_ef = PPA_GRID_OPTIONS[st.selectbox("Kanıtlı Elektrik Faktörü", list(PPA_GRID_OPTIONS.keys()))] if has_ppa else EU_DEFAULT_TR_GRID_EF

                prec_cn_idx = list(CATALOG.keys()).index(vd["prec_cn"]) if vd["prec_cn"] in CATALOG else 0
                prec_cn = st.selectbox("Öncül Hammadde CN", list(CATALOG.keys()), index=prec_cn_idx, format_func=lambda x: f"{format_cn(x)} ({CATALOG[x]['s']})")
                pc1, pc2 = st.columns(2)
                prec_mi = pc1.number_input("Öncül Oranı m_i (t/t)", min_value=0.0, value=float(vd["prec_mi"]), step=0.05)
                prec_origin = pc2.selectbox("Öncül Menşei", ["Türkiye / Diğer Üçüncü Ülke", "AB / EFTA Menşeli", "Hurda / Sıfır Gömülü Emisyon", "Bilinmeyen Menşe (Annex IV En Yüksek)"])
                prec_verified = st.checkbox("Tedarikçi Verisi Doğrulanmış", value=bool(vd["prec_verified"]))
                prec_actual_see = st.number_input("Doğrulanmış Öncül SEE (t/t)", min_value=0.0, value=float(vd["prec_actual_see"]), step=0.05)
                prec_route = st.selectbox("Öncül Üretim Rotası", ["(E)", "(C)", "(D)", "(K)", "(L)", "GENEL"])

            run_btn = st.form_submit_button("🛡️ İhracat Ceza Riskini Hesapla & AB Alıcı Raporu Oluştur", use_container_width=True, type="primary")

        # Bilişsel Yük Yönetimi: Tesis Sicil, White-Label Logo & Şifreli Kasa Kapalı Çekmecede
        with st.expander("⚙️ Gelişmiş Tesis Sicil, White-Label Logo & Şifreli Kasa (.cbamvault)", expanded=False):
            broker_name = st.text_input("Danışmanlık / Kurum Unvanı (White-Label)", value="TR-CBAM Kurumsal Uyum")
            broker_logo = st.file_uploader("Kurumsal Logo Yükle (PDF Anteti)", type=["png", "jpg", "jpeg"])
            installation_id = st.text_input("CBAM Registry Installation ID", value=vd["installation_id"])
            unlocode = st.text_input("UN/LOCODE", value=vd["unlocode"])
            coordinates = st.text_input("Coğrafi Koordinatlar", value=vd["coordinates"])
            verifier_id = st.text_input("Akredite Doğrulayıcı ID ((EU) 2025/2547)", value="PENDING-VERIFIER-2026")

            st.markdown("---")
            vault_pass = st.text_input("Tesis Kasa Parolası (AES-256)", type="password", value="")
            uploaded_vault = st.file_uploader("Kayıtlı Kasa Yükle (.cbamvault)", type=["cbamvault"])
            if uploaded_vault and vault_pass:
                if st.button("🔓 Kasayı Çöz ve Yükle", use_container_width=True):
                    try:
                        loaded_dict = decrypt_vault_payload(uploaded_vault.getvalue(), vault_pass)
                        st.session_state["vault_data"].update(loaded_dict)
                        vd = st.session_state["vault_data"]
                        st.success(f"Kasa yüklendi: {vd['company_name']}")
                    except Exception:
                        st.error("Parola hatalı veya kasa dosyası bozuk.")
        st.markdown('</div>', unsafe_allow_html=True)

    # Varsayılan başlangıç hesaplaması (ilk açılışta da ekran dolu ve canlı görünsün)
    if not run_btn and "cached_run" not in st.session_state:
        fuels_input = {"dogalgaz": float(vd["q_gas"]), "tas_komuru": 0.0, "kok_komuru": 0.0, "linyit": 0.0, "fuel_oil": 0.0, "lpg": 0.0}
        precursors_input = [{"cn": vd["prec_cn"], "m_i": float(vd["prec_mi"]), "origin": "Türkiye / Diğer Üçüncü Ülke", "verified": False, "actual_see": 0.45, "route": "(E)"}]
        res = calculate_cbam_core(
            vd["selected_cn"], vd["selected_route"], float(vd["eu_tonnage"]), float(vd["total_prod_ton"]),
            fuels_input, 0.0, float(vd["elec_kwh"]), EU_DEFAULT_TR_GRID_EF,
            precursors_input, 1.0, cbam_cert_price, local_cp_eff, yr_rule
        )
        meta_obj = {
            "company_name": vd["company_name"], "buyer_company": vd["buyer_company"],
            "broker_name": broker_name, "installation_id": installation_id,
            "unlocode": unlocode, "coordinates": coordinates, "verifier_id": verifier_id,
            "tonnage": float(vd["eu_tonnage"]), "invoice_status": invoice_status, "year": yr_rule["year"]
        }
        findings, score, crit_c, warn_c, pass_c, bm_thermo, elec_pct, fuel_pct = run_verifier_audit_matrix(
            vd["selected_cn"], res["sector"], vd["selected_route"], float(vd["eu_tonnage"]), float(vd["total_prod_ton"]),
            fuels_input, 0.0, float(vd["elec_kwh"]), False, precursors_input, 1.0, res, meta_obj
        )
        st.session_state["last_audit_findings"] = findings
    else:
        if run_btn:
            fuels_input = {"dogalgaz": q_gas, "tas_komuru": q_coal, "kok_komuru": q_coke, "linyit": q_lignite, "fuel_oil": q_fueloil, "lpg": q_lpg}
            precursors_input = [{"cn": prec_cn, "m_i": prec_mi, "origin": prec_origin, "verified": prec_verified, "actual_see": prec_actual_see, "route": prec_route}]
            st.session_state["vault_data"].update({
                "company_name": company_name, "buyer_company": buyer_company,
                "installation_id": installation_id, "unlocode": unlocode, "coordinates": coordinates,
                "selected_cn": selected_cn, "selected_route": selected_route,
                "total_prod_ton": total_prod_ton, "eu_tonnage": eu_tonnage,
                "q_gas": q_gas, "elec_kwh": elec_kwh,
                "prec_cn": prec_cn, "prec_mi": prec_mi,
                "prec_verified": prec_verified, "prec_actual_see": prec_actual_see,
            })
            st.session_state["cached_run"] = {
                "cn": selected_cn, "route": selected_route, "eu_t": eu_tonnage, "tot_t": total_prod_ton,
                "fuels": fuels_input, "proc": process_em_ton, "elec": elec_kwh, "grid": active_grid_ef,
                "precs": precursors_input, "func": func_ratio, "has_ppa": has_ppa,
                "comp": company_name, "buyer": buyer_company, "inv": invoice_status
            }

        cr = st.session_state["cached_run"]
        res = calculate_cbam_core(
            cr["cn"], cr["route"], cr["eu_t"], cr["tot_t"],
            cr["fuels"], cr["proc"], cr["elec"], cr["grid"],
            cr["precs"], cr["func"], cbam_cert_price, local_cp_eff, yr_rule
        )
        meta_obj = {
            "company_name": cr["comp"], "buyer_company": cr["buyer"],
            "broker_name": broker_name, "installation_id": installation_id,
            "unlocode": unlocode, "coordinates": coordinates, "verifier_id": verifier_id,
            "tonnage": cr["eu_t"], "invoice_status": cr["inv"], "year": yr_rule["year"]
        }
        findings, score, crit_c, warn_c, pass_c, bm_thermo, elec_pct, fuel_pct = run_verifier_audit_matrix(
            cr["cn"], res["sector"], cr["route"], cr["eu_t"], cr["tot_t"],
            cr["fuels"], cr["proc"], cr["elec"], cr["has_ppa"], cr["precs"], cr["func"], res, meta_obj
        )
        st.session_state["last_audit_findings"] = findings

    with col_right:
        # ÜST SATIR: 3 MODÜLER FIORI KPI KARTI
        k1, k2, k3 = st.columns(3)
        with k1:
            st.markdown(f"""
            <div class="fiori-tile fiori-tile-danger">
                <div class="tile-header">AB Varsayılan Tarife Cezası</div>
                <div class="tile-bottom-row">
                    <div class="tile-kpi" style="color:var(--danger);">{res['tax_default']:,.0f} €</div>
                    <div class="tile-pill" style="background:var(--danger-bg); color:var(--danger);">+%{res['markup_pct']} Marj</div>
                </div>
                <div class="tile-footnote">AB alıcısının faturanızdan kesebileceği tutar</div>
            </div>
            """, unsafe_allow_html=True)
        with k2:
            st.markdown(f"""
            <div class="fiori-tile fiori-tile-neutral">
                <div class="tile-header">Tesis Gerçek SKDM Yükü</div>
                <div class="tile-bottom-row">
                    <div class="tile-kpi" style="color:var(--text-primary);">{res['tax_actual']:,.0f} €</div>
                    <div class="tile-pill" style="background:rgba(148,163,184,0.16); color:var(--text-primary);">{res['actual_see_per_ton']:.3f} t/t</div>
                </div>
                <div class="tile-footnote">Col A Kıyas Değeri: {res['col_a_bm']:.3f} tCO₂e/t</div>
            </div>
            """, unsafe_allow_html=True)
        with k3:
            st.markdown(f"""
            <div class="fiori-tile fiori-tile-success">
                <div class="tile-header">Korunan Net İhracat Avantajı</div>
                <div class="tile-bottom-row">
                    <div class="tile-kpi" style="color:var(--success); font-size:1.85rem;">+{res['direct_savings']:,.0f} €</div>
                    <div class="tile-pill" style="background:var(--success-bg); color:var(--success);">+{res['savings_per_ton_eur']:,.1f} €/t</div>
                </div>
                <div class="tile-footnote">AB alıcısına sunulan net maliyet avantajı</div>
            </div>
            """, unsafe_allow_html=True)

        st.write("")
        # ORTA SATIR: %65 PLOTLY MALİYET KIYASLAMA + %35 TÜV/SGS UYUM ENDEKSİ
        mid_left, mid_right = st.columns([1.85, 1.0])
        with mid_left:
            st.markdown('<div class="fiori-tile">', unsafe_allow_html=True)
            fig = go.Figure(data=[
                go.Bar(
                    x=["Kanıtsız Default Ceza (A)", "Tesis Gerçek Yükü (B)"],
                    y=[res["tax_default"], res["tax_actual"]],
                    marker_color=[active_theme["danger"], active_theme["success"]],
                    text=[f"{res['tax_default']:,.0f} €", f"{res['tax_actual']:,.0f} €"],
                    textposition="auto",
                    width=[0.42, 0.42]
                )
            ])
            fig.update_layout(
                template="none",
                title=dict(text=f"{meta_obj['company_name']} — {yr_rule['year']} SKDM Maliyet Kıyaslaması (€)", font=dict(size=14)),
                height=250,
                margin=dict(l=16, r=16, t=40, b=16),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color=active_theme["text_primary"], family=selected_font_name),
                xaxis=dict(showgrid=False, linecolor=active_theme["border"]),
                yaxis=dict(showgrid=True, gridcolor="rgba(148,163,184,0.18)")
            )
            st.plotly_chart(fig, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with mid_right:
            score_color = active_theme["success"] if score >= 80 else (active_theme["accent"] if score >= 60 else active_theme["danger"])
            st.markdown(f"""
            <div class="fiori-tile fiori-tile-neutral">
                <div class="tile-header">TÜV / SGS Denetim Skoru</div>
                <div class="tile-bottom-row">
                    <div class="tile-kpi" style="color:{score_color};">%{score} / 100</div>
                    <div class="tile-pill" style="background:rgba(148,163,184,0.15); color:var(--text-primary);">{crit_c} Kritik</div>
                </div>
                <div class="tile-footnote" style="margin-top:14px; line-height:1.55;">
                    <strong>Türkiye Termodinamik Kıyas:</strong><br/>
                    • Elektrik: <strong>P{elec_pct}</strong> ({res['elec_mwh_per_ton']:.2f} MWh/t)<br/>
                    • Yakıt: <strong>P{fuel_pct}</strong> ({res['fuel_gj_per_ton']:.2f} GJ/t)<br/>
                    • Rota: <strong>{res['route_tag']}</strong> ({res['sector']})
                </div>
            </div>
            """, unsafe_allow_html=True)

        # ALT SATIR: HIZLI AKSİYON ŞERİDİ (ACTION PILLS & ENDOWMENT PAYWALL)
        logo_b = broker_logo.getvalue() if broker_logo else None
        pdf_out = generate_corporate_pdf_core(meta_obj, res, {"score": score, "crit": crit_c, "warn": warn_c, "elec_pct": elec_pct, "fuel_pct": fuel_pct}, is_licensed=is_licensed, logo_bytes=logo_b)

        st.write("")
        if is_licensed:
            buyer_pdf_out = generate_eu_buyer_pitch_pdf(meta_obj, res, meta_obj["buyer_company"])
            eu_xlsx_out = generate_eu_communication_template_xlsx(meta_obj, res)
            xml_row = dict(res)
            xml_row["tonnage"] = meta_obj["tonnage"]
            xml_out = generate_cbam_xml_core(meta_obj, [xml_row])

            a1, a2, a3, a4 = st.columns(4)
            a1.download_button("📥 Kurumsal Rapor PDF", data=pdf_out, file_name=f"TR_CBAM_Rapor_{res['cn_raw']}.pdf", mime="application/pdf", use_container_width=True)
            a2.download_button("🌍 AB Alıcı Teklifi EN PDF", data=buyer_pdf_out, file_name=f"EU_CBAM_Buyer_Offer_{res['cn_raw']}.pdf", mime="application/pdf", use_container_width=True)
            a3.download_button("📗 AB İletişim Şablonu .xlsx", data=eu_xlsx_out, file_name=f"EU_CBAM_Comm_Template_{res['cn_raw']}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            a4.download_button("🧾 CBAM Registry .xml", data=xml_out, file_name=f"EU_CBAM_O3CI_{res['cn_raw']}.xml", mime="application/xml", use_container_width=True)
        else:
            a1, a2, a3, a4 = st.columns(4)
            a1.download_button("📥 Taslak Rapor PDF", data=pdf_out, file_name=f"TASLAK_TR_CBAM_{res['cn_raw']}.pdf", mime="application/pdf", use_container_width=True)
            a2.button("🔒 AB Alıcı Teklifi EN PDF", disabled=True, use_container_width=True)
            a3.button("🔒 AB İletişim Şablonu .xlsx", disabled=True, use_container_width=True)
            a4.button("🔒 CBAM Registry .xml", disabled=True, use_container_width=True)

            st.markdown(f"""
            <div class="fiori-paywall">
                <div style="font-weight:700; font-size:0.92rem; color:var(--text-primary);">🔒 Kurumsal Çıktı Paketi Kilitli (Deneme Modu)</div>
                <div style="font-size:0.82rem; color:var(--text-secondary); margin-top:4px;">
                    <strong>{meta_obj['company_name']}</strong> adına hesaplanan <strong>+{res['direct_savings']:,.0f} €</strong> maliyet avantajını belgeleyen
                    filigransız İngilizce AB Alıcı Teklif PDF'ini, Resmi AB Excel İletişim Şablonunu (<code>.xlsx</code>) ve O3CI <code>.xml</code> dosyasını
                    indirmek için sağ üstteki <strong>Profil & Lisans</strong> menüsünden kurumsal anahtarınızı girin.
                </div>
                <div class="blur-doc-preview">
                    <div><strong>📄 EU_CBAM_Buyer_Offer_{res['cn_raw']}.pdf</strong> — Prepared by: {meta_obj['company_name']} | For: {meta_obj['buyer_company']} | Saving: +{res['direct_savings']:,.2f} EUR</div>
                    <div class="blur-overlay">Installation ID: {installation_id} | Actual SEE: {res['actual_see_per_ton']:.4f} tCO2e/t | Column A Benchmark: {res['col_a_bm']:.3f} | O3CI v3.2 XML Verified Payload...</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        if vault_pass:
            vault_bytes = encrypt_vault_payload(st.session_state["vault_data"], vault_pass)
            st.write("")
            st.download_button("🔐 Şifreli Tesis Kasasını İndir (.cbamvault)", data=vault_bytes, file_name=f"{meta_obj['company_name'].replace(' ', '_')}_{yr_rule['year']}.cbamvault", mime="application/json", use_container_width=True)

# ==============================================================================
# SEKME 2: GÜMRÜK MÜŞAVİRİ REFERRAL
# ==============================================================================
with tab2:
    st.markdown('<div class="fiori-tile">', unsafe_allow_html=True)
    st.subheader("📨 Gümrük Müşaviri Sıfır-Sorumluluk Müşteri Yönlendirme (Referral)")
    st.caption("Gümrük müşavirliklerinin müteselsil sorumluluk veya teknik danışmanlık yükü almadan ihracatçı müşterilerine 2026 SKDM Risk Özeti iletmesini sağlar.")

    gm1, gm2 = st.columns(2, gap="large")
    with gm1:
        gm_client_name = st.text_input("İhracatçı Müşteri Unvanı", value="Kocaeli Metal Profil Sanayi A.Ş.")
        gm_dec_no = st.text_input("AB İhracat Beyanname / Referans No", value="26340100EX004812")
        gm_cn = st.selectbox("Beyannamedeki Ana GTİP Kodu", list(CATALOG.keys()), index=list(CATALOG.keys()).index("72085120") if "72085120" in CATALOG else 0, format_func=lambda x: f"{format_cn(x)} ({CATALOG[x]['s']})")
        gm_ton = st.number_input("Sevkiyat Net Ağırlığı (Ton)", min_value=1.0, value=85.0, step=5.0)

    gm_item = CATALOG[gm_cn]
    est_def_tax = gm_item["dv"] * yr_rule["std_markup"] * gm_ton * cbam_cert_price
    est_save = est_def_tax * 0.62

    with gm2:
        g_k1, g_k2 = st.columns(2)
        g_k1.markdown(f'<div class="fiori-tile fiori-tile-danger"><div class="tile-header">Cezalı Default Riski</div><div class="tile-kpi" style="color:var(--danger);">{est_def_tax:,.0f} €</div></div>', unsafe_allow_html=True)
        g_k2.markdown(f'<div class="fiori-tile fiori-tile-success"><div class="tile-header">Tahmini Kurtarılan</div><div class="tile-kpi" style="color:var(--success);">+{est_save:,.0f} €</div></div>', unsafe_allow_html=True)
        st.write("")
        invite_msg = f"""Sayın {gm_client_name} Dış Ticaret / Finans Yönetimi,

Müşavirliğimiz aracılığıyla tescil edilen {gm_dec_no} referanslı ({format_cn(gm_cn)} GTİP kodlu, {gm_ton:,.1f} Ton) AB ihracat sevkiyatınız, 2026 AB Sınırda Karbon Düzenlemesi (SKDM) kesin dönem kapsamındadır.

Tesis enerji veriniz AB ithalatçısına iletilmediği takdirde, AB gümrüğünde varsayılan ceza tarifesi üzerinden yaklaşık {est_def_tax:,.0f} € sertifika maliyeti oluşmakta ve ithalatçı müşteriniz bu farkı faturanızdan kesme/sipariş iptali hakkı kazanmaktadır.

TR-CBAM platformu üzerinden dönemlik elektrik ve doğalgaz tüketiminizi girerek bu maliyetin yaklaşık {est_save:,.0f} € tutarındaki kısmını sıfırlayabilir ve AB alıcınıza özel teklif belgesi oluşturabilirsiniz.

Saygılarımızla,
{broker_name}"""
        st.text_area("Müşteriye Gönderilecek Hazır Bilgilendirme Metni", value=invite_msg, height=215)
    st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# SEKME 3: ERP VERİ AKTARIMI
# ==============================================================================
with tab3:
    st.markdown('<div class="fiori-tile">', unsafe_allow_html=True)
    st.subheader("🔌 Logo Tiger / Mikro / Netsis / SAP Hızlı Veri Aktarım Modülü")
    erp_type = st.selectbox("Kaynak ERP / Muhasebe Sistemi", ["Logo Tiger / Go3", "Mikro Yazılım", "Netsis", "SAP S/4HANA (MM/PP)", "Diğer Özel Excel / CSV"])
    erp_file = st.file_uploader(f"{erp_type} Dışa Aktarım Dosyasını Yükle (.xlsx / .csv)", type=["xlsx", "csv"])

    if erp_file:
        try:
            df_erp = pd.read_csv(erp_file) if erp_file.name.lower().endswith(".csv") else pd.read_excel(erp_file)
            st.dataframe(df_erp.head(5), use_container_width=True)
            cols = list(df_erp.columns)

            m_c1, m_c2, m_c3, m_c4 = st.columns(4)
            col_prod = m_c1.selectbox("Toplam Üretim Miktarı Sütunu (Ton)", cols)
            col_eu = m_c2.selectbox("AB İhracat Miktarı Sütunu (Ton)", cols)
            col_elec = m_c3.selectbox("Elektrik Tüketimi Sütunu (kWh)", cols)
            col_gas = m_c4.selectbox("Doğalgaz Tüketimi Sütunu (m³)", cols)

            if st.button("⚡ ERP Verilerini Topla ve Ana Hesaplama Motoruna Aktar", type="primary", use_container_width=True):
                st.session_state["vault_data"]["total_prod_ton"] = float(pd.to_numeric(df_erp[col_prod], errors="coerce").sum())
                st.session_state["vault_data"]["eu_tonnage"] = float(pd.to_numeric(df_erp[col_eu], errors="coerce").sum())
                st.session_state["vault_data"]["elec_kwh"] = float(pd.to_numeric(df_erp[col_elec], errors="coerce").sum())
                st.session_state["vault_data"]["q_gas"] = float(pd.to_numeric(df_erp[col_gas], errors="coerce").sum())
                st.success(f"{erp_type} verileri başarıyla toplandı ve 1. Sekmeye aktarıldı!")
        except Exception as e:
            st.error(f"ERP dosyası okunamadı: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# SEKME 4: TEDARİKÇİ (PRECURSOR) PORTALI
# ==============================================================================
with tab4:
    st.markdown('<div class="fiori-tile">', unsafe_allow_html=True)
    st.subheader("🤝 Tedarikçi Emisyon Veri Portalı (Precursor Network Effect)")
    t_col1, t_col2 = st.columns(2, gap="large")
    with t_col1:
        st.markdown("#### 1. Tedarikçiye Veri Talebi Oluştur")
        sup_name = st.text_input("Tedarikçi Firma Unvanı", value="Anadolu Haddecilik / Biyet Sanayi A.Ş.")
        sup_cn = st.selectbox("Tedarik Edilen Öncül Hammadde CN Kodu", list(CATALOG.keys()), index=list(CATALOG.keys()).index("72071114") if "72071114" in CATALOG else 0, format_func=lambda x: f"{format_cn(x)} - {CATALOG[x]['d'][:50]}")
        sup_period = st.selectbox("Talep Edilen Veri Dönemi", ["2026 Q1–Q4 (Tam Yıl)", "2026 Q3", "2026 Q4"])

        req_letter = f"""Sayın {sup_name} Yetkilisi,

Avrupa Birliği Sınırda Karbon Düzenleme Mekanizması (SKDM / CBAM) Kesin Dönem Tüzüğü ((EU) 2023/956 ve (EU) 2025/2547) kapsamında, firmanızdan tedarik ettiğimiz {format_cn(sup_cn)} GTİP kodlu öncül hammaddenin spesifik gömülü emisyon (SEE) verisini AB gümrük beyanımıza eklememiz gerekmektedir.

Tedarikçi verisi beyan edilmediği takdirde AB gümrüklerinde {CATALOG[sup_cn]['dv']} tCO2e/ton baz değerine ek +%10 ceza marjı ({CATALOG[sup_cn]['dv']*1.10:.3f} tCO2e/ton) uygulanmaktadır.

Lütfen aşağıdaki 3 bilgiyi doldurarak tarafımıza iletiniz:
1. Tesis Üretim Rotası: (Örn: Hurda/EAF, Yüksek Fırın BF/BOF, İkincil Ergitme)
2. Dönem ({sup_period}) Spesifik Doğrudan Emisyon (SEE_direct - tCO2e/ton):
3. Akredite Doğrulayıcı ((EU) 2025/2547) Durumu ve Belge No:

Saygılarımızla,
{vd['company_name']} — Dış Ticaret & Sürdürülebilirlik Birimi"""
        st.text_area("Tedarikçiye Gönderilecek Resmi Talep Metni", value=req_letter, height=250)

    with t_col2:
        st.markdown("#### 2. Tedarikçi Verisini Sisteme Aktar")
        in_sup_cn = st.selectbox("Gelen Tedarikçi CN Kodu", list(CATALOG.keys()), index=list(CATALOG.keys()).index("72071114") if "72071114" in CATALOG else 0, format_func=lambda x: f"{format_cn(x)} ({CATALOG[x]['s']})")
        in_sup_see = st.number_input("Tedarikçi Beyan Edilen SEE (tCO₂e / Ton)", min_value=0.01, value=0.42, step=0.02)
        in_sup_mi = st.number_input("1 Ton Mamuldeki Kullanım Oranı m_i (t/t)", min_value=1.00, value=1.05, step=0.01)
        in_sup_ver = st.checkbox("Tedarikçi Verisi Akredite Doğrulayıcı Onaylı", value=True)

        if st.button("⚡ Tedarikçi Verisini Ana Hesaplama Motoruna Aktar", use_container_width=True, type="primary"):
            st.session_state["vault_data"]["prec_cn"] = in_sup_cn
            st.session_state["vault_data"]["prec_actual_see"] = in_sup_see
            st.session_state["vault_data"]["prec_mi"] = in_sup_mi
            st.session_state["vault_data"]["prec_verified"] = in_sup_ver
            st.success("Tedarikçi verisi ana hesaplama sekmesine aktarıldı!")
    st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# SEKME 5: ÇOKLU SEVKİYAT & 50 TON EŞİĞİ
# ==============================================================================
with tab5:
    st.markdown('<div class="fiori-tile">', unsafe_allow_html=True)
    st.subheader("📊 Çoklu Sevkiyat Analizi & (EU) 2025/2083 Madde 2a Kümülatif 50 Ton Kontrolü")
    sample_df = pd.DataFrame([
        {"Urun_Adi": "Sıcak Haddelenmiş Sac", "CN_Kodu": "72085120", "Rota": "(E)", "Tesis_Uretim_Ton": 1000.0, "AB_Sevkiyat_Ton": 30.0, "Dogalgaz_m3": 18000.0, "Elektrik_kWh": 150000.0, "Oncul_CN": "72071114", "Oncul_mi": 1.05},
        {"Urun_Adi": "Çelik Cıvata", "CN_Kodu": "73181588", "Rota": "(E)", "Tesis_Uretim_Ton": 400.0, "AB_Sevkiyat_Ton": 15.0, "Dogalgaz_m3": 6000.0, "Elektrik_kWh": 80000.0, "Oncul_CN": "72131000", "Oncul_mi": 1.08},
        {"Urun_Adi": "Alüminyum Profil", "CN_Kodu": "76041090", "Rota": "(L)", "Tesis_Uretim_Ton": 500.0, "AB_Sevkiyat_Ton": 20.0, "Dogalgaz_m3": 9000.0, "Elektrik_kWh": 110000.0, "Oncul_CN": "76011090", "Oncul_mi": 1.04},
    ])
    xbuf = io.BytesIO()
    with pd.ExcelWriter(xbuf, engine="openpyxl") as w:
        sample_df.to_excel(w, index=False, sheet_name="Sevkiyatlar")

    st.download_button("📄 Örnek Çoklu Sevkiyat Şablonunu İndir (.xlsx)", data=xbuf.getvalue(), file_name="TR_CBAM_Toplu_Sablon.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    up_xlsx = st.file_uploader("Doldurulmuş Sevkiyat Excel Dosyasını Yükle", type=["xlsx"])
    if up_xlsx:
        try:
            df_in = pd.read_excel(up_xlsx)
            batch_res, table_rows = [], []
            cum_mass = 0.0

            for _, row in df_in.iterrows():
                c_raw = re.sub(r"\s+", "", str(row["CN_Kodu"]).strip())
                if c_raw not in CATALOG:
                    c_raw = next((k for k in CATALOG if k.startswith(c_raw[:4])), "72085120")
                rt = str(row.get("Rota", "(E)")).strip()
                eu_t = float(row["AB_Sevkiyat_Ton"])
                tot_t = float(row.get("Tesis_Uretim_Ton", eu_t))
                if c_raw != "28041000":
                    cum_mass += eu_t

                p_cn_raw = re.sub(r"\s+", "", str(row.get("Oncul_CN", "72071114")).strip())
                if p_cn_raw not in CATALOG:
                    p_cn_raw = "72071114"

                r = calculate_cbam_core(
                    c_raw, rt, eu_t, tot_t,
                    {"dogalgaz": float(row.get("Dogalgaz_m3", 0.0))}, 0.0,
                    float(row.get("Elektrik_kWh", 0.0)), EU_DEFAULT_TR_GRID_EF,
                    [{"cn": p_cn_raw, "m_i": float(row.get("Oncul_mi", 0.0)), "origin": "Türkiye / Diğer Üçüncü Ülke", "verified": False, "actual_see": 0.0, "route": rt}],
                    1.0, cbam_cert_price, local_cp_eff, yr_rule
                )
                r["tonnage"] = eu_t
                batch_res.append(r)
                table_rows.append({
                    "Ürün": row.get("Urun_Adi", r["cn"]),
                    "CN Kodu": r["cn"], "Rota": rt, "AB Tonaj": eu_t,
                    "Gerçek SEE (t/t)": round(r["actual_see_per_ton"], 3),
                    "Kanıtsız Ceza (€)": round(r["tax_default"], 2),
                    "Kanıtlı Vergi (€)": round(r["tax_actual"], 2),
                    "Net Tasarruf (€)": round(r["direct_savings"], 2),
                })

            if cum_mass < 50.0:
                st.success(f"✅ **Madde 2a Kümülatif Eşik Muafiyeti:** Yıllık kümülatif miktar **{cum_mass:.2f} Ton** (< 50 Ton) olduğu için CBAM yükümlülüğü doğmaz.")
            else:
                st.warning(f"⚠️ **Madde 2a Kümülatif Eşik Aşıldı:** Kümülatif toplam **{cum_mass:.2f} Ton** (≥ 50 Ton). Tam CBAM sertifika yükümlülüğü geçerlidir.")

            df_out = pd.DataFrame(table_rows)
            st.dataframe(df_out, use_container_width=True)

            if is_licensed:
                meta_batch = {
                    "company_name": "Toplu İhracat Beyanı", "broker_name": broker_name,
                    "installation_id": installation_id, "unlocode": unlocode,
                    "coordinates": coordinates, "verifier_id": verifier_id, "year": yr_rule["year"]
                }
                b_xml = generate_cbam_xml_core(meta_batch, batch_res)
                st.download_button("🧾 Toplu Beyan İçin O3CI v3.2 XML İndir", data=b_xml, file_name=f"EU_CBAM_Batch_O3CI_{yr_rule['year']}.xml", mime="application/xml", use_container_width=True)
            else:
                st.info("🔒 Toplu beyan XML çıktısı kurumsal lisans anahtarı ile aktifleşir.")
        except Exception as e:
            st.error(f"Excel işleme hatası: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# SEKME 6: RESMİ KATALOG (570 CN KODU)
# ==============================================================================
with tab6:
    st.markdown('<div class="fiori-tile">', unsafe_allow_html=True)
    st.subheader("📚 Avrupa Komisyonu Resmi 2026 Kesin Dönem Türkiye & Benchmark Kataloğu (570 CN Kodu)")
    cat_rows = []
    for cn_8, v in CATALOG.items():
        cat_rows.append({
            "CN Kodu": format_cn(cn_8),
            "Sektör": v["s"],
            "Ürün Tanımı": v["d"],
            "Kaynak": "Türkiye (2026/1740)" if v["src"] == "TR" else "Other Countries Fallback",
            "2026 Baz Default (tCO₂e/t)": v["dv"],
            "Annex IV En Yüksek (tCO₂e/t)": v["a4"],
            "Column A Benchmark (Proses)": ", ".join([f"{x[0]} {x[1]}".strip() for x in v["a"]]),
            "Column B Benchmark (Tam Zincir)": ", ".join([f"{x[0]} {x[1]}".strip() for x in v["b"]]),
        })
    st.dataframe(pd.DataFrame(cat_rows), use_container_width=True, height=550)
    st.markdown('</div>', unsafe_allow_html=True)
