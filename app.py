import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from fpdf import FPDF
from pypdf import PdfReader
from PIL import Image
import openpyxl
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime, timezone
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
    page_title="TR-CBAM v4.0 | AB SKDM Kesin Dönem Kokpiti",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- SABİT MEVZUAT KATSAYILARI ((EU) 2025/2547 Ek II Bölüm G & ETKB) ---
FUEL_FACTORS = {
    "dogalgaz": {"name": "Doğalgaz (m³)", "ef": 0.00192, "unit": "m³"},
    "tas_komuru": {"name": "Taş Kömürü (Ton)", "ef": 2.420, "unit": "Ton"},
    "kok_komuru": {"name": "Kok Kömürü (Ton)", "ef": 3.100, "unit": "Ton"},
    "linyit": {"name": "Linyit Kömürü (Ton)", "ef": 1.150, "unit": "Ton"},
    "fuel_oil": {"name": "Fuel-Oil (Ton)", "ef": 3.120, "unit": "Ton"},
    "lpg": {"name": "LPG (Ton)", "ef": 2.980, "unit": "Ton"},
}

# Resmi AB Türkiye Varsayılan Elektrik Şebeke Katsayısı ((EU) 2025/2621 Annex II)
EU_DEFAULT_TR_GRID_EF = 0.718  # tCO2e / MWh

# ETKB Tüketim Noktası Katsayıları (Yalnızca PPA / Doğrudan Teknik Bağlantı Kanıtlıysa - (EU) 2025/2547 Ek II D.4)
PPA_GRID_OPTIONS = {
    "ETKB Dağıtım Hattı / OSB Bağlantısı (0,469 tCO₂e/MWh)": 0.469,
    "ETKB İletim Hattı / Yüksek Gerilim (0,436 tCO₂e/MWh)": 0.436,
    "Sıfır Emisyonlu Yenilenebilir PPA / GES (0,000 tCO₂e/MWh)": 0.000,
}

# Üretim Rotası Açıklamaları ((EU) 2025/2620 Ek Bölüm 5.3)
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

# (EU) 2023/956 Madde 31 Faz Takvimi & (EU) 2025/2621 Ceza Marjları
YEAR_RULES = {
    "2026 (Kesin Dönem 1. Yıl | Ceza: +%10 [Gübre +%1] | CBAM Faktörü: %97,5)": {
        "std_markup": 1.10, "fert_markup": 1.01, "cbam_factor": 0.975, "year": 2026, "period_code": "(1)"
    },
    "2027 (Kesin Dönem 2. Yıl | Ceza: +%20 [Gübre +%1] | CBAM Faktörü: %95,0)": {
        "std_markup": 1.20, "fert_markup": 1.01, "cbam_factor": 0.950, "year": 2027, "period_code": "(1)"
    },
    "2028 (Kesin Dönem 3. Yıl | Ceza: +%30 [Gübre +%1] | CBAM Faktörü: %90,0)": {
        "std_markup": 1.30, "fert_markup": 1.01, "cbam_factor": 0.900, "year": 2028, "period_code": "(2)"
    },
    "2030 (Orta Vade | Ceza: +%30 [Gübre +%1] | CBAM Faktörü: %51,5)": {
        "std_markup": 1.30, "fert_markup": 1.01, "cbam_factor": 0.515, "year": 2030, "period_code": "(2)"
    },
    "2034 (Tam Uygulama | Ceza: +%30 [Gübre +%1] | CBAM Faktörü: %0,0)": {
        "std_markup": 1.30, "fert_markup": 1.01, "cbam_factor": 0.000, "year": 2034, "period_code": "(2)"
    },
}

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
        curr_sec = None
        for r in range(3, ws.max_row + 1):
            cn_raw = ws.cell(r, 1).value
            desc = ws.cell(r, 2).value
            if cn_raw is None:
                continue
            cn_str = str(cn_raw).strip()
            if cn_str in ["Cement", "Fertilisers", "Aluminium", "Hydrogen", "Iron and steel", "Iron & Steel"]:
                curr_sec = cn_str
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
        return _parse_excel_files("CBAM Benchmarks_20260206.xlsx", "DV correcting act_final update_06.08.xlsx"), "Resmi Excel Dosyaları (570 CN Kodu)"
    if os.path.exists("cbam_catalog_b64.txt"):
        with open("cbam_catalog_b64.txt", "r", encoding="utf-8") as f:
            raw = gzip.decompress(base64.b64decode(f.read().strip())).decode("utf-8")
        return json.loads(raw), "Gömülü Resmi AB Veri Tabanı (570 CN Kodu)"
    # Yedek minimal fallback
    return {
        "72085120": {"d": "Flat-rolled products of iron or non-alloy steel, hot-rolled", "s": "Demir-Çelik", "dv": 2.428, "dv_d": 2.428, "dv_i": 0.0, "src": "TR", "dr": "(C)", "a4": 8.23, "a": [[0.044, ""]], "b": [[1.37, "(C)"], [0.481, "(D)"], [0.072, "(E)"]]}
    }, "Acil Yedek Mod"

CATALOG, CATALOG_SOURCE = load_full_catalog()

def format_cn(cn_8):
    if len(cn_8) == 8:
        return f"{cn_8[:4]} {cn_8[4:6]} {cn_8[6:8]}"
    return cn_8

def extract_routes_for_item(item):
    routes = []
    for _, ind in item.get("b", []) + item.get("a", []):
        m = re.findall(r"\(([A-L])\)", ind)
        for code in m:
            tag = f"({code})"
            if tag not in routes:
                routes.append(tag)
    if not routes:
        routes = ["GENEL"]
    return routes

def select_benchmark_value(bm_list, route_tag, period_code):
    if not bm_list:
        return 0.0
    # 1. Tam eşleşme: hem rota hem dönem (1) veya (2)
    for val, ind in bm_list:
        if (route_tag == "GENEL" or route_tag in ind) and period_code in ind:
            return float(val)
    # 2. Rota eşleşmesi (dönem bağımsız)
    for val, ind in bm_list:
        if route_tag != "GENEL" and route_tag in ind:
            return float(val)
    # 3. Dönem eşleşmesi (rota bağımsız)
    for val, ind in bm_list:
        if period_code in ind:
            return float(val)
    # 4. İlk satır (sabit değer)
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

# --- (EU) 2025/2547 & (EU) 2025/2620 & (EU) 2026/2498 TAM HESAPLAMA MOTORU ---
def calculate_cbam_v4(
    cn_code, route_tag, eu_tonnage, total_prod_ton,
    fuels_dict, process_em_ton, elec_kwh, grid_ef,
    precursors_list, func_ratio,
    cbam_cert_price, local_cp_eff, yr_rule
):
    item = CATALOG[cn_code]
    sector = item["s"]
    period_code = yr_rule["period_code"]
    cbam_factor = yr_rule["cbam_factor"]

    # Ceza marjı (Gübre +%1 sabit, diğerleri +%10 / +%20 / +%30)
    markup = yr_rule["fert_markup"] if sector == "Gübre" else yr_rule["std_markup"]
    markup_pct = int(round((markup - 1.0) * 100))

    # Fonksiyonel birim düzeltmesi (Çimento: klinker oranı | Gübre: azot oranı | Diğerleri: 1.0)
    effective_func_ratio = max(0.01, min(1.0, func_ratio))

    # Üretim paylaştırma (Annex IV Point 2)
    eff_total_prod = max(eu_tonnage, total_prod_ton)
    eu_share = eu_tonnage / eff_total_prod

    # 1. Tesis Doğrudan Proses Emisyonu (ae_g - Yakıtlar + Proses Emisyonları)
    total_fuel_em = sum(qty * FUEL_FACTORS[fkey]["ef"] for fkey, qty in fuels_dict.items())
    total_facility_direct = total_fuel_em + process_em_ton
    ae_g_per_ton = total_facility_direct / eff_total_prod  # 1 ton ürün başına kendi tesis doğrudan emisyonu

    # Dolaylı emisyon (Sadece Çimento ve Gübrede mali yükümlülüğe girer)
    indirect_per_ton = ((elec_kwh / 1000.0) * grid_ef) / eff_total_prod
    is_indirect_billable = sector in ["Çimento", "Gübre"]

    # 2. Öncül Maddeler (Precursors) — (EU) 2025/2547 Denklem 61 & (EU) 2025/2620 Denklem 4
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
                # Doğrulanmış öncül: kendi rotasının Column B (veya A) kıyas değeri
                bm_p = select_benchmark_value(p_item["b"], p_route, period_code)
                sefa_i = cbam_factor * 1.0 * bm_p
            else:
                # Doğrulanmamış öncül: Ülke Default + Mark-up (veya Bilinmeyen Menşe Annex IV en yüksek)
                if p_origin == "Bilinmeyen Menşe (Annex IV En Yüksek Ceza)":
                    base_p_dv = p_item["a4"]
                else:
                    base_p_dv = p_item["dv"]
                p_markup = yr_rule["fert_markup"] if p_item["s"] == "Gübre" else yr_rule["std_markup"]
                see_i = base_p_dv * p_markup
                def_p_route = p_item["dr"] if p_item["dr"] else p_route
                bm_p = select_benchmark_value(p_item["b"], def_p_route, period_code)
                sefa_i = cbam_factor * 1.0 * bm_p

        prec_see_sum_per_ton += m_i * see_i
        prec_sefa_sum_per_ton += m_i * sefa_i
        prec_details.append({
            "cn": format_cn(p_cn),
            "m_i": m_i,
            "origin": p_origin,
            "see_i": see_i,
            "sefa_i": sefa_i
        })

    # 3. Senaryo B (Gerçek Tesis Verisi) Toplam Gömülü Emisyon & Ücretsiz Tahsisat ((EU) 2025/2620 Denklem 2–4)
    actual_see_per_ton = (ae_g_per_ton + (indirect_per_ton if is_indirect_billable else 0.0) + prec_see_sum_per_ton) * effective_func_ratio
    actual_total_em = actual_see_per_ton * eu_tonnage

    # Column A (BM_g*) Proses Kıyas Değeri + Öncül SEFA Toplamı
    col_a_bm = select_benchmark_value(item["a"], route_tag, period_code)
    sfa_proc_per_ton = cbam_factor * 1.0 * col_a_bm
    sefa_actual_per_ton = (sfa_proc_per_ton + prec_sefa_sum_per_ton) * effective_func_ratio
    faa_actual_total = sefa_actual_per_ton * eu_tonnage

    # 4. Senaryo A (Kanıtsız Default Beyan) & Column B (BM_g) Ücretsiz Tahsisatı ((EU) 2025/2620 Denklem 6)
    default_ef_with_markup = item["dv"] * markup * effective_func_ratio
    default_total_em = default_ef_with_markup * eu_tonnage

    default_route = item["dr"] if item["dr"] else route_tag
    col_b_bm = select_benchmark_value(item["b"], default_route, period_code)
    sefa_default_per_ton = cbam_factor * 1.0 * col_b_bm * effective_func_ratio
    faa_default_total = sefa_default_per_ton * eu_tonnage

    # 5. Net Teslim Edilecek Sertifika Sayısı
    cert_default = max(0.0, default_total_em - faa_default_total)
    cert_actual = max(0.0, actual_total_em - faa_actual_total)

    # 6. (EU) 2026/2498 Madde 7 Resmi Karbon Fiyatı Mahsubu
    # Reduction = (EFF_CP / RefPrice_CBAM) * Q_g  => Net Ödenecek = (Sertifika - Reduction) * RefPrice_CBAM
    cp_reduction_default = min(cert_default, (local_cp_eff / cbam_cert_price) * cert_default) if cbam_cert_price > 0 else 0.0
    cp_reduction_actual = min(cert_actual, (local_cp_eff / cbam_cert_price) * cert_actual) if cbam_cert_price > 0 else 0.0

    net_cert_default = max(0.0, cert_default - cp_reduction_default)
    net_cert_actual = max(0.0, cert_actual - cp_reduction_actual)

    tax_default = net_cert_default * cbam_cert_price
    tax_actual = net_cert_actual * cbam_cert_price
    direct_savings = max(0.0, tax_default - tax_actual)

    # 2027+ Çeyreklik %50 Sertifika Bulundurma Yükümlülüğü ((EU) 2025/2083 Madde 22)
    quarterly_holding_certs = net_cert_actual * 0.50 if yr_rule["year"] >= 2027 else 0.0

    return {
        "cn": format_cn(cn_code),
        "cn_raw": cn_code,
        "desc": item["d"],
        "sector": sector,
        "dv_source": "Türkiye ((EU) 2025/2621 & 2026/1740)" if item["src"] == "TR" else "Diğer Ülkeler Fallback ((EU) 2025/2621)",
        "route_tag": route_tag,
        "default_route": default_route,
        "markup_pct": markup_pct,
        "cbam_factor": cbam_factor,
        "func_ratio": effective_func_ratio,
        "eff_total_prod": eff_total_prod,
        "eu_share_pct": eu_share * 100.0,
        "ae_g_per_ton": ae_g_per_ton,
        "indirect_per_ton": indirect_per_ton,
        "is_indirect_billable": is_indirect_billable,
        "prec_see_per_ton": prec_see_sum_per_ton,
        "prec_sefa_per_ton": prec_sefa_sum_per_ton,
        "prec_details": prec_details,
        "actual_see_per_ton": actual_see_per_ton,
        "actual_total_em": actual_total_em,
        "default_ef_markup": default_ef_with_markup,
        "default_total_em": default_total_em,
        "col_a_bm": col_a_bm,
        "col_b_bm": col_b_bm,
        "faa_actual_total": faa_actual_total,
        "faa_default_total": faa_default_total,
        "cert_default": cert_default,
        "cert_actual": cert_actual,
        "cp_reduction_actual": cp_reduction_actual,
        "net_cert_default": net_cert_default,
        "net_cert_actual": net_cert_actual,
        "tax_default": tax_default,
        "tax_actual": tax_actual,
        "direct_savings": direct_savings,
        "quarterly_holding_certs": quarterly_holding_certs,
        "is_de_minimis_candidate": (eu_tonnage < 50.0) and (cn_code != "28041000"),
    }

# --- O3CI v3.2 & (EU) 2025/2547 EK IV UYUMLU XML ÜRETİCİ ---
def generate_cbam_xml_v4(meta, results_list):
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

# --- KURUMSAL PDF RAPOR ÜRETİCİ ---
def setup_unicode_font(pdf):
    if os.path.exists("DejaVuSans.ttf"):
        pdf.add_font("TrFont", "", "DejaVuSans.ttf")
        pdf.add_font("TrFont", "B", "DejaVuSans-Bold.ttf" if os.path.exists("DejaVuSans-Bold.ttf") else "DejaVuSans.ttf")
        pdf.add_font("TrFont", "I", "DejaVuSans-Oblique.ttf" if os.path.exists("DejaVuSans-Oblique.ttf") else "DejaVuSans.ttf")
        return "TrFont"
    return "Helvetica"

def generate_corporate_pdf_v4(meta, res, logo_bytes=None):
    pdf = FPDF()
    pdf.add_page()
    font_name = setup_unicode_font(pdf)

    temp_logo = None
    if logo_bytes:
        try:
            img = Image.open(io.BytesIO(logo_bytes))
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                img.save(tmp.name, format="PNG")
                temp_logo = tmp.name
            pdf.image(temp_logo, x=12, y=10, w=24)
            pdf.set_xy(38, 12)
        except Exception:
            pdf.set_xy(10, 12)
    else:
        pdf.set_xy(10, 12)

    pdf.set_fill_color(15, 23, 42)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(font_name, "B", 10.8)
    pdf.cell(0, 10.5, "TR-CBAM v4.0 | AB SKDM KESİN DÖNEM DENETİM ÖNCESİ RAPORU", ln=True, align="C", fill=True)

    pdf.set_text_color(80, 80, 80)
    pdf.set_font(font_name, "", 7.8)
    pdf.cell(0, 5.5, f"Tarih: {datetime.now().strftime('%d.%m.%Y %H:%M')}  |  Tesis ID: {meta['installation_id']}  |  UN/LOCODE: {meta['unlocode']}  |  Yıl: {meta['year']}", ln=True, align="C")
    pdf.ln(3)

    pdf.set_text_color(0, 0, 0)
    pdf.set_font(font_name, "B", 9.4)
    pdf.cell(0, 6, "1. TESİS KİMLİĞİ, ÜRETİM ROTASI VE FONKSİYONEL BİRİM KAPSAMI", ln=True)
    pdf.set_font(font_name, "", 8.3)
    pdf.cell(0, 4.8, f"İhracatçı Tesis: {meta['company_name']}  |  Koordinat: {meta['coordinates']}", ln=True)
    pdf.cell(0, 4.8, f"CN Kodu: {res['cn']} — {res['desc'][:68]}", ln=True)
    pdf.cell(0, 4.8, f"Sektör: {res['sector']}  |  Seçilen Üretim Rotası: {ROUTE_LABELS.get(res['route_tag'], res['route_tag'])}", ln=True)
    pdf.cell(0, 4.8, f"Toplam Tesis Üretimi: {res['eff_total_prod']:,.2f} t  |  AB Sevkiyatı: {meta['tonnage']:,.2f} t (AB Payı: %{res['eu_share_pct']:.1f})  |  Fonk. Oran: %{res['func_ratio']*100:.1f}", ln=True)
    pdf.cell(0, 4.8, f"Default Kaynağı: {res['dv_source']}  |  Enerji Kanıtı: {meta['invoice_status']}", ln=True)
    pdf.ln(2.5)

    pdf.set_fill_color(241, 245, 249)
    pdf.set_font(font_name, "B", 9.4)
    pdf.cell(0, 7, "2. AB SKDM SERTİFİKA YÜKÜMLÜLÜĞÜ KIYASLAMASI ((EU) 2025/2620 & 2026/2498)", ln=True, fill=True)
    pdf.ln(1.2)

    pdf.set_font(font_name, "B", 7.8)
    pdf.cell(72, 6.5, "Beyan Senaryosu", border=1)
    pdf.cell(38, 6.5, "Gömülü Emisyon", border=1, align="C")
    pdf.cell(40, 6.5, "Net Sertifika (Adet)", border=1, align="C")
    pdf.cell(40, 6.5, "Ödenecek Tutar", border=1, align="R", ln=True)

    pdf.set_font(font_name, "", 7.8)
    pdf.cell(72, 6.5, f"A: Kanıtsız Default (+%{res['markup_pct']} Marj | Col B)", border=1)
    pdf.cell(38, 6.5, f"{res['default_total_em']:,.2f} tCO2e", border=1, align="C")
    pdf.cell(40, 6.5, f"{res['net_cert_default']:,.2f}", border=1, align="C")
    pdf.cell(40, 6.5, f"{res['tax_default']:,.2f} EUR", border=1, align="R", ln=True)

    pdf.cell(72, 6.5, "B: Kanıtlı Gerçek Veri (Col A + Precursor)", border=1)
    pdf.cell(38, 6.5, f"{res['actual_total_em']:,.2f} tCO2e", border=1, align="C")
    pdf.cell(40, 6.5, f"{res['net_cert_actual']:,.2f}", border=1, align="C")
    pdf.cell(40, 6.5, f"{res['tax_actual']:,.2f} EUR", border=1, align="R", ln=True)

    pdf.set_font(font_name, "B", 9.0)
    pdf.set_fill_color(220, 252, 231)
    pdf.cell(150, 8, "DOĞRUDAN KORUNAN NET SKDM SERTİFİKA TASARRUFU (A - B)", border=1, fill=True)
    pdf.cell(40, 8, f"+{res['direct_savings']:,.2f} EUR", border=1, align="R", ln=True, fill=True)
    pdf.ln(2.5)

    pdf.set_font(font_name, "B", 9.4)
    pdf.cell(0, 6, "3. ANNEX III EMİSYON KIRILIMI & MADDE 31 BENCHMARK DETAYI", ln=True)
    pdf.set_font(font_name, "", 8.2)
    pdf.cell(0, 4.6, f"• Tesis Doğrudan Proses Emisyonu (ae_g): {res['ae_g_per_ton']:.4f} tCO2e/t", ln=True)
    pdf.cell(0, 4.6, f"• Dolaylı Elektrik Emisyonu: {res['indirect_per_ton']:.4f} tCO2e/t ({'Hesaba Dahil' if res['is_indirect_billable'] else 'Annex II Kapsam Dışı'})", ln=True)
    pdf.cell(0, 4.6, f"• Öncül Hammadde Toplam Gömülü Emisyonu (∑ m_i x SEE_i): {res['prec_see_per_ton']:.4f} tCO2e/t", ln=True)
    pdf.cell(0, 4.6, f"• Gerçek Veri Ücretsiz Tahsisat Düşümü (Col A BM={res['col_a_bm']:.3f} + Precursor SEFA): -{res['faa_actual_total']:,.2f} tCO2e", ln=True)
    pdf.cell(0, 4.6, f"• Kanıtsız Beyan Ücretsiz Tahsisat Düşümü (Col B BM={res['col_b_bm']:.3f}): -{res['faa_default_total']:,.2f} tCO2e", ln=True)
    if res["cp_reduction_actual"] > 0:
        pdf.cell(0, 4.6, f"• (EU) 2026/2498 Madde 7 Üçüncü Ülke Karbon Fiyatı Mahsubu: -{res['cp_reduction_actual']:,.2f} Sertifika", ln=True)
    pdf.ln(2)

    pdf.set_font(font_name, "B", 8.3)
    pdf.cell(0, 5, f"Yetkili Gümrük Müşavirliği / Beyan Temsilcisi: {meta['broker_name']}  |  Doğrulayıcı ID: {meta['verifier_id']}", ln=True)
    pdf.ln(1)
    pdf.set_font(font_name, "", 7.0)
    pdf.set_text_color(75, 75, 75)
    pdf.multi_cell(0, 3.6, "Mevzuat Uyum Beyanı: Bu rapor (EU) 2023/956, (EU) 2025/2083, (EU) 2025/2547 (Ek II–IV), (EU) 2025/2620 (Madde 2–4 Column A/B Kıyas Değerleri), (EU) 2025/2621, (EU) 2026/1740 ve (EU) 2026/2498 sayılı AB SKDM tüzüklerine uygun olarak hesaplanmıştır. Zero-Retention politikası gereği veriler kalıcı olarak saklanmaz.")

    out_bytes = bytes(pdf.output())
    if temp_logo and os.path.exists(temp_logo):
        try:
            os.remove(temp_logo)
        except Exception:
            pass
    return out_bytes

# --- YAN MENÜ (TESİS KİMLİĞİ & PİYASA / MEVZUAT PARAMETRELERİ) ---
with st.sidebar:
    st.header("🏢 Tesis & O3CI v3.2 Kayıt Bilgileri")
    broker_name = st.text_input("Gümrük Müşavirliği / Temsilci Unvanı", value="İÜ Girişim Hub — R&D Division")
    broker_logo = st.file_uploader("Kurumsal Logo Yükle (PDF Anteti)", type=["png", "jpg", "jpeg"])
    installation_id = st.text_input("CBAM Registry Installation ID", value="TR-INST-2026-0001")
    unlocode = st.text_input("UN/LOCODE (Şehir/Liman Kodu)", value="TRIST")
    coordinates = st.text_input("Tesis Coğrafi Koordinatları (Enlem, Boylam)", value="41.0082, 28.9784")
    verifier_id = st.text_input("Akredite Doğrulayıcı ID ((EU) 2025/2547)", value="PENDING-VERIFIER-2026")

    st.divider()
    st.subheader("📅 Dönem, Piyasa & TR-ETS Kuralları")
    year_choice = st.selectbox("Beyan Yılı & Faz Takvimi", list(YEAR_RULES.keys()))
    yr_rule = YEAR_RULES[year_choice]

    cbam_cert_price = st.number_input(
        "AB CBAM Sertifika Fiyatı (€ / tCO₂e)",
        min_value=10.0, value=82.32, step=0.5,
        help="2026 yılında çeyreklik ortalama EU ETS ihale fiyatı; 2027'den itibaren haftalık ortalama fiyat uygulanır."
    )

    # Madde 7: TR-ETS 2026-2027 Pilot Dönem Kısıtı
    if yr_rule["year"] in [2026, 2027]:
        st.warning("🇹🇷 **TR-ETS Pilot Dönem Kuralı:** 2026 (Raporlama) ve 2027 (%100 Ücretsiz Tahsisat) yıllarında fiili karbon ödemesi bulunmadığından (EU) 2026/2498 kapsamında mahsup **0,00 €/tCO₂e** olarak kilitlidir.")
        local_cp_eff = 0.0
    else:
        local_cp_eff = st.number_input(
            "TR-ETS Fiilen Ödenen Efektif Karbon Fiyatı (€ / tCO₂e)",
            min_value=0.0, value=15.0, step=1.0,
            help="(EU) 2026/2498 Madde 7 uyarınca 2028+ uygulama döneminde ödenen net karbon fiyatı mahsup edilir."
        )

    st.divider()
    st.caption(f"📦 **Veri Tabanı:** {CATALOG_SOURCE}")
    st.caption("🔒 **Zero-Retention:** Tüm veriler yalnızca aktif oturum RAM belleğinde işlenir.")

# --- ANA EKRAN ---
st.title("TR-CBAM v4.0 | AB SKDM Kesin Dönem Uyum & Mühendislik Kokpiti")
st.markdown("**570 Resmi CN Kodu ((EU) 2026/1740) • Column A vs. B Kıyas Motoru ((EU) 2025/2620) • Öncül Madde ($m_i$) & TR-ETS ((EU) 2026/2498)**")

tab1, tab2, tab3 = st.tabs([
    "🏭 Tekil Sevkiyat & Mühendislik Hesabı",
    "📊 Çoklu Sevkiyat (Excel) & Kümülatif 50 Ton Kontrolü",
    "📚 Tam Resmi Katalog (570 CN Kodu & Benchmarklar)"
])

# ==============================================================================
# SEKME 1: TEKİL SEVKİYAT ANALİZİ
# ==============================================================================
with tab1:
    col_left, col_right = st.columns([1.2, 1])

    with col_left:
        st.subheader("1. Ürün, Üretim Rotası ve Tesis Girdileri")

        # 570 CN Kodu İçinde Arama & Seçim
        search_q = st.text_input("🔍 GTİP / CN Kodu veya Ürün Adı Ara (Örn: 7208, 7318, 7604, 2523, 3102, bolt, aluminium)", value="")
        filtered_cns = [
            cn for cn, info in CATALOG.items()
            if not search_q.strip() or search_q.lower() in cn.lower() or search_q.lower() in format_cn(cn).lower() or search_q.lower() in info["d"].lower()
        ]
        if not filtered_cns:
            filtered_cns = list(CATALOG.keys())

        default_idx = filtered_cns.index("72085120") if "72085120" in filtered_cns else 0
        selected_cn = st.selectbox(
            f"CN / GTİP Kodu Seçin ({len(filtered_cns)} sonuç)",
            filtered_cns,
            index=default_idx,
            format_func=lambda x: f"{format_cn(x)} | [{CATALOG[x]['s']}] {CATALOG[x]['d'][:75]}"
        )
        sel_item = CATALOG[selected_cn]

        # Üretim Rotası Seçici
        avail_routes = extract_routes_for_item(sel_item)
        default_route_idx = 0
        if sel_item["s"] == "Demir-Çelik" and "(E)" in avail_routes:
            default_route_idx = avail_routes.index("(E)")
        elif sel_item["s"] == "Alüminyum" and "(L)" in avail_routes:
            default_route_idx = avail_routes.index("(L)")

        selected_route = st.selectbox(
            "Üretim Rotası ((EU) 2025/2620 Ek Bölüm 5.3)",
            avail_routes,
            index=default_route_idx,
            format_func=lambda r: ROUTE_LABELS.get(r, r)
        )

        # Fonksiyonel Birim Oranı (Çimento Klinker / Gübre Azot)
        func_ratio = 1.0
        if sel_item["s"] == "Çimento" and not selected_cn.startswith("252310"):
            clinker_pct = st.slider("Çimento İçindeki Klinker Oranı (%) — (EU) 2025/2547 Ek III Denklem 64", 10, 100, 75)
            func_ratio = clinker_pct / 100.0
        elif sel_item["s"] == "Gübre":
            nitrogen_pct = st.slider("Gübre Azot (N) İçeriği (%) — (EU) 2025/2547 Ek III Denklem 66", 5, 82, 27)
            func_ratio = nitrogen_pct / 100.0

        # Fatura OCR
        invoice_file = st.file_uploader("Opsiyonel: Fatura Yükle (PDF / JPG / PNG - Otomatik Okuma)", type=["pdf", "png", "jpg", "jpeg"])
        auto_kwh, auto_gas = 180000.0, 24000.0
        invoice_status = "Kanıt Belgesi Yüklenmedi (Ön Simülasyon)"
        if invoice_file:
            dk, dg = extract_energy_from_file(invoice_file)
            invoice_status = f"Doğrulandı: {invoice_file.name}"
            if dk:
                auto_kwh = dk
                st.success(f"Faturadan otomatik elektrik okundu: **{dk:,.2f} kWh**")
            if dg:
                auto_gas = dg
                st.success(f"Faturadan otomatik doğalgaz okundu: **{dg:,.2f} m³**")

        with st.form("v4_main_form"):
            company_name = st.text_input("İhracatçı Tesis Unvanı", value="Örnek Metal ve Çelik Sanayi A.Ş.")
            p1, p2 = st.columns(2)
            total_prod_ton = p1.number_input("Tesis Dönemlik Toplam Üretimi (Ton)", min_value=1.0, value=1000.0, step=50.0)
            eu_tonnage = p2.number_input("AB İhracat Sevkiyatı (Ton)", min_value=0.1, value=120.0, step=10.0)

            st.markdown("**🔥 Yakıt Tüketimleri & Proses Emisyonları (Dönemlik Tesis Toplamı — (EU) 2025/2547 Ek II)**")
            f1, f2, f3 = st.columns(3)
            q_gas = f1.number_input("Doğalgaz (m³)", min_value=0.0, value=float(auto_gas), step=500.0)
            q_coal = f2.number_input("Taş Kömürü (Ton)", min_value=0.0, value=0.0, step=1.0)
            q_coke = f3.number_input("Kok Kömürü (Ton)", min_value=0.0, value=0.0, step=1.0)

            f4, f5, f6 = st.columns(3)
            q_lignite = f4.number_input("Linyit (Ton)", min_value=0.0, value=0.0, step=1.0)
            q_fueloil = f5.number_input("Fuel-Oil (Ton)", min_value=0.0, value=0.0, step=1.0)
            q_lpg = f6.number_input("LPG (Ton)", min_value=0.0, value=0.0, step=1.0)

            process_em_ton = st.number_input(
                "Doğrudan Proses Emisyonları (tCO₂e — Kalsinasyon, EAF Elektrot, Karbonat Ayrışması)",
                min_value=0.0, value=0.0, step=5.0
            )

            st.markdown("**⚡ Elektrik Tüketimi & Şebeke Kuralı ((EU) 2025/2547 Ek II Bölüm D.4)**")
            e1, e2 = st.columns([1, 1.3])
            elec_kwh = e1.number_input("Tesis Elektrik Tüketimi (kWh)", min_value=0.0, value=float(auto_kwh), step=1000.0)
            has_ppa = e2.checkbox("Fiziksel PPA veya Doğrudan Teknik Bağlantı + Saatlik Sayaç Kanıtlı", value=False)
            if has_ppa:
                ppa_choice = st.selectbox("Kanıtlı Elektrik Faktörü", list(PPA_GRID_OPTIONS.keys()))
                active_grid_ef = PPA_GRID_OPTIONS[ppa_choice]
            else:
                active_grid_ef = EU_DEFAULT_TR_GRID_EF
                st.caption(f"ℹ️ PPA/Doğrudan bağlantı işaretli değil: Resmi AB Türkiye Şebeke Faktörü (**{EU_DEFAULT_TR_GRID_EF} tCO₂e/MWh**) uygulanır.")

            st.markdown("**🧱 Öncül Hammadde (Precursor) Girdileri — Complex Goods ((EU) 2025/2547 Denklem 61)**")
            p_col1, p_col2, p_col3 = st.columns(3)
            prec_cn = p_col1.selectbox(
                "Öncül Hammadde CN Kodu",
                list(CATALOG.keys()),
                index=list(CATALOG.keys()).index("72071114") if "72071114" in CATALOG else 0,
                format_func=lambda x: f"{format_cn(x)} ({CATALOG[x]['s']})"
            )
            prec_mi = p_col2.number_input("Tüketim Oranı m_i (Ton Öncül / Ton Ürün)", min_value=0.0, value=1.05 if sel_item["s"] in ["Demir-Çelik", "Alüminyum"] else 0.0, step=0.05)
            prec_origin = p_col3.selectbox("Öncül Menşei", ["Türkiye / Diğer Üçüncü Ülke", "AB / EFTA Menşeli", "Hurda / Sıfır Gömülü Emisyon", "Bilinmeyen Menşe (Annex IV En Yüksek)"])

            pv1, pv2, pv3 = st.columns(3)
            prec_verified = pv1.checkbox("Tedarikçi Verisi Akredite Doğrulanmış", value=False)
            prec_actual_see = pv2.number_input("Doğrulanmış Öncül SEE (tCO₂e/t)", min_value=0.0, value=0.45, step=0.05)
            prec_route = pv3.selectbox("Öncül Üretim Rotası", ["(E)", "(C)", "(D)", "(K)", "(L)", "GENEL"])

            run_btn = st.form_submit_button("Kesin Dönem Analizini & Belgeleri Üret (v4.0)", use_container_width=True, type="primary")

    if run_btn:
        fuels_input = {
            "dogalgaz": q_gas, "tas_komuru": q_coal, "kok_komuru": q_coke,
            "linyit": q_lignite, "fuel_oil": q_fueloil, "lpg": q_lpg
        }
        precursors_input = [{
            "cn": prec_cn, "m_i": prec_mi, "origin": prec_origin,
            "verified": prec_verified, "actual_see": prec_actual_see, "route": prec_route
        }]
        res = calculate_cbam_v4(
            selected_cn, selected_route, eu_tonnage, total_prod_ton,
            fuels_input, process_em_ton, elec_kwh, active_grid_ef,
            precursors_input, func_ratio, cbam_cert_price, local_cp_eff, yr_rule
        )

        with col_right:
            st.subheader("2. Finansal & Mevzuat Sonuçları")
            if res["is_de_minimis_candidate"]:
                st.warning("⚠️ **Madde 2a De Minimis Notu:** AB ithalatçısının **tüm tedarikçilerden yıllık kümülatif ithalatı** 50 tonun altındaysa muafiyet uygulanır (Hidrojen ve Elektrik hariçtir).")
            if res["quarterly_holding_certs"] > 0:
                st.info(f"📅 **2027+ Çeyreklik Sertifika Bulundurma Yükümlülüğü ((EU) 2025/2083 Madde 22):** Her çeyrek sonunda en az **{res['quarterly_holding_certs']:,.2f} adet** (%50) CBAM sertifikası hesapta bulundurulmalıdır.")

            m1, m2, m3 = st.columns(3)
            m1.metric("Kanıtsız Ceza (A)", f"{res['tax_default']:,.0f} €", help=f"Col B BM: {res['col_b_bm']:.3f} | Net Sertifika: {res['net_cert_default']:,.2f}")
            m2.metric("Kanıtlı Beyan (B)", f"{res['tax_actual']:,.0f} €", help=f"Col A BM: {res['col_a_bm']:.3f} | SEE: {res['actual_see_per_ton']:.3f} t/t")
            m3.metric("Net Vergi Tasarrufu", f"+{res['direct_savings']:,.0f} €")

            fig = go.Figure(data=[
                go.Bar(
                    x=["A: Kanıtsız Default (Col B)", "B: Kanıtlı Tesis + Öncül (Col A)"],
                    y=[res["tax_default"], res["tax_actual"]],
                    marker_color=["#ef4444", "#10b981"],
                    text=[f"{res['tax_default']:,.0f} €", f"{res['tax_actual']:,.0f} €"],
                    textposition="auto"
                )
            ])
            fig.update_layout(
                title=f"{yr_rule['year']} Yılı Net AB SKDM Sertifika Maliyeti (€)",
                height=300, margin=dict(l=20, r=20, t=40, b=20), yaxis_title="Euro (€)"
            )
            st.plotly_chart(fig, use_container_width=True)

            st.markdown("#### 🔍 Mühendislik ve Benchmark Kırılımı")
            st.write(f"- **Tesis Doğrudan Emisyonu ($ae_g$):** `{res['ae_g_per_ton']:.4f} tCO₂e/t`")
            st.write(f"- **Öncül Hammadde Emisyonu ($\\sum m_i \\times SEE_i$):** `{res['prec_see_per_ton']:.4f} tCO₂e/t`")
            st.write(f"- **Toplam Gerçek Gömülü Emisyon ($SEE_g$):** `{res['actual_see_per_ton']:.4f} tCO₂e/t`")
            st.write(f"- **Senaryo B Ücretsiz Tahsisat (Col A `{res['col_a_bm']:.3f}` + Precursor SEFA):** `-{res['faa_actual_total']:,.2f} tCO₂e`")
            st.write(f"- **Senaryo A Ücretsiz Tahsisat (Col B `{res['col_b_bm']:.3f}`):** `-{res['faa_default_total']:,.2f} tCO₂e`")

        st.divider()
        b1, b2 = st.columns(2)
        meta_v4 = {
            "company_name": company_name, "broker_name": broker_name,
            "installation_id": installation_id, "unlocode": unlocode,
            "coordinates": coordinates, "verifier_id": verifier_id,
            "tonnage": eu_tonnage, "invoice_status": invoice_status,
            "year": yr_rule["year"]
        }
        logo_b = broker_logo.getvalue() if broker_logo else None
        pdf_out = generate_corporate_pdf_v4(meta_v4, res, logo_b)

        xml_row = dict(res)
        xml_row["tonnage"] = eu_tonnage
        xml_out = generate_cbam_xml_v4(meta_v4, [xml_row])

        b1.download_button(
            "📥 Kurumsal Pre-Audit PDF Raporunu İndir (v4.0)",
            data=pdf_out, file_name=f"TR_CBAM_v4_{yr_rule['year']}_{selected_cn}.pdf",
            mime="application/pdf", use_container_width=True
        )
        b2.download_button(
            "🧾 O3CI v3.2 / AB CBAM Registry XML İndir",
            data=xml_out, file_name=f"EU_CBAM_O3CI_v32_{selected_cn}_{yr_rule['year']}.xml",
            mime="application/xml", use_container_width=True
        )

# ==============================================================================
# SEKME 2: ÇOKLU KALEM & KÜMÜLATİF 50 TON EŞİK KONTROLÜ
# ==============================================================================
with tab2:
    st.subheader("Çoklu Sevkiyat Analizi & (EU) 2025/2083 Madde 2a Kümülatif 50 Ton Kontrolü")
    sample_df = pd.DataFrame([
        {"Urun_Adi": "Sıcak Haddelenmiş Sac", "CN_Kodu": "72085120", "Rota": "(E)", "Tesis_Uretim_Ton": 1000.0, "AB_Sevkiyat_Ton": 30.0, "Dogalgaz_m3": 18000.0, "Elektrik_kWh": 150000.0, "Oncul_CN": "72071114", "Oncul_mi": 1.05},
        {"Urun_Adi": "Çelik Cıvata", "CN_Kodu": "73181588", "Rota": "(E)", "Tesis_Uretim_Ton": 400.0, "AB_Sevkiyat_Ton": 15.0, "Dogalgaz_m3": 6000.0, "Elektrik_kWh": 80000.0, "Oncul_CN": "72131000", "Oncul_mi": 1.08},
        {"Urun_Adi": "Alüminyum Profil", "CN_Kodu": "76041090", "Rota": "(L)", "Tesis_Uretim_Ton": 500.0, "AB_Sevkiyat_Ton": 20.0, "Dogalgaz_m3": 9000.0, "Elektrik_kWh": 110000.0, "Oncul_CN": "76011090", "Oncul_mi": 1.04},
    ])
    xbuf = io.BytesIO()
    with pd.ExcelWriter(xbuf, engine="openpyxl") as w:
        sample_df.to_excel(w, index=False, sheet_name="Sevkiyatlar")

    st.download_button(
        "📄 Örnek Çoklu Sevkiyat Şablonunu İndir (.xlsx)",
        data=xbuf.getvalue(), file_name="TR_CBAM_v4_Toplu_Sablon.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    up_xlsx = st.file_uploader("Doldurulmuş Sevkiyat Excel Dosyasını Yükle", type=["xlsx"])
    if up_xlsx:
        try:
            df_in = pd.read_excel(up_xlsx)
            batch_res = []
            table_rows = []
            cum_mass_subject_to_threshold = 0.0

            for _, row in df_in.iterrows():
                c_raw = re.sub(r"\s+", "", str(row["CN_Kodu"]).strip())
                if c_raw not in CATALOG:
                    c_raw = next((k for k in CATALOG if k.startswith(c_raw[:4])), "72085120")
                rt = str(row.get("Rota", "(E)")).strip()
                eu_t = float(row["AB_Sevkiyat_Ton"])
                tot_t = float(row.get("Tesis_Uretim_Ton", eu_t))
                if c_raw != "28041000":
                    cum_mass_subject_to_threshold += eu_t

                p_cn_raw = re.sub(r"\s+", "", str(row.get("Oncul_CN", "72071114")).strip())
                if p_cn_raw not in CATALOG:
                    p_cn_raw = "72071114"
                p_mi_val = float(row.get("Oncul_mi", 0.0))

                r = calculate_cbam_v4(
                    c_raw, rt, eu_t, tot_t,
                    {"dogalgaz": float(row.get("Dogalgaz_m3", 0.0))}, 0.0,
                    float(row.get("Elektrik_kWh", 0.0)), EU_DEFAULT_TR_GRID_EF,
                    [{"cn": p_cn_raw, "m_i": p_mi_val, "origin": "Türkiye / Diğer Üçüncü Ülke", "verified": False, "actual_see": 0.0, "route": rt}],
                    1.0, cbam_cert_price, local_cp_eff, yr_rule
                )
                r["tonnage"] = eu_t
                batch_res.append(r)
                table_rows.append({
                    "Ürün": row.get("Urun_Adi", r["cn"]),
                    "CN Kodu": r["cn"],
                    "Rota": rt,
                    "AB Tonaj": eu_t,
                    "Gerçek SEE (t/t)": round(r["actual_see_per_ton"], 3),
                    "Kanıtsız Ceza (€)": round(r["tax_default"], 2),
                    "Kanıtlı Vergi (€)": round(r["tax_actual"], 2),
                    "Net Tasarruf (€)": round(r["direct_savings"], 2),
                })

            if cum_mass_subject_to_threshold < 50.0:
                st.success(f"✅ **Madde 2a Kümülatif Eşik Muafiyeti:** İthalatçının yıllık toplam kümülatif miktarı **{cum_mass_subject_to_threshold:.2f} Ton** (< 50 Ton) olduğu için CBAM sertifika yükümlülüğünden muaftır.")
            else:
                st.warning(f"⚠️ **Madde 2a Kümülatif Eşik Aşıldı:** Kümülatif toplam **{cum_mass_subject_to_threshold:.2f} Ton** (≥ 50 Ton). Tüm kalemler tam CBAM yükümlülüğüne tabidir.")

            df_out = pd.DataFrame(table_rows)
            st.dataframe(df_out, use_container_width=True)
            st.metric("Toplam Sevkiyat Tasarrufu", f"+{df_out['Net Tasarruf (€)'].sum():,.2f} €")

            meta_batch = {
                "company_name": "Toplu İhracat Beyanı", "broker_name": broker_name,
                "installation_id": installation_id, "unlocode": unlocode,
                "coordinates": coordinates, "verifier_id": verifier_id,
                "year": yr_rule["year"]
            }
            b_xml = generate_cbam_xml_v4(meta_batch, batch_res)
            st.download_button(
                "🧾 Toplu Beyan İçin O3CI v3.2 XML İndir",
                data=b_xml, file_name=f"EU_CBAM_Batch_O3CI_v32_{yr_rule['year']}.xml",
                mime="application/xml", use_container_width=True
            )
        except Exception as e:
            st.error(f"Excel işleme hatası: {e}")

# ==============================================================================
# SEKME 3: 570 CN KODLU TAM RESMİ KATALOG
# ==============================================================================
with tab3:
    st.subheader("Avrupa Komisyonu Resmi 2026 Kesin Dönem Türkiye & Benchmark Kataloğu (570 CN Kodu)")
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
