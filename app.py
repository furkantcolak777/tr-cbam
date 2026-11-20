import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from fpdf import FPDF
from pypdf import PdfReader
from PIL import Image
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime, timezone
import tempfile
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
    page_title="TR-CBAM v3.0 | Resmi Mevzuat Uyum Kokpiti",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- RESMİ EMİSYON VE BENCHMARK VERİ TABANI ---
# 1. T.C. ETKB Tüketim Noktası Ulusal Elektrik Şebekesi Emisyon Faktörleri (tCO2e / MWh)
ETKB_GRID_FACTORS = {
    "ETKB Dağıtım Hattı / OSB Bağlantısı (0,469 tCO₂e/MWh)": 0.469,
    "ETKB İletim Hattı / Yüksek Gerilim (0,436 tCO₂e/MWh)": 0.436,
    "ETKB Türkiye Geneli Üretim Ortalaması (0,434 tCO₂e/MWh)": 0.434,
    "AB CBAM Türkiye Varsayılan Şebeke Faktörü (0,718 tCO₂e/MWh)": 0.718,
}

# 2. AB Uygulama Tüzüğü (EU) 2023/1773 Doğalgaz Yanma Faktörü (56,1 tCO2/TJ * 0,0342 TJ/1000m3)
NAT_GAS_EF = 0.00192  # tCO2 / m3

# 3. (EU) 2025/2621 & (EU) 2026/1740 Türkiye Özel Varsayılan Değerleri + (EU) 2025/2620 B Sütunu Resmi Benchmarklar
# markup_type: "standard" (2026: +%10, 2027: +%20, 2028+: +%30) | "fertiliser" (2026-2028+: +%1 sabit)
GTIP_2026_DB = {
    "7207 11 14 - Yarı Mamul Alaşımsız Çelik (Kütük / Slab)": {
        "cn": "7207 11 14", "base_default": 2.310, "eu_benchmark": 1.364, "precursor_ef": 1.450,
        "scope": "direct_only", "sector": "Demir-Çelik", "markup_type": "standard"
    },
    "7208 51 20 - Sıcak Haddelenmiş Yassı Çelik Ürünler": {
        "cn": "7208 51 20", "base_default": 2.428, "eu_benchmark": 1.370, "precursor_ef": 1.350,
        "scope": "direct_only", "sector": "Demir-Çelik", "markup_type": "standard"
    },
    "7209 - Soğuk Haddelenmiş Yassı Çelik Ürünler": {
        "cn": "7209", "base_default": 2.580, "eu_benchmark": 1.458, "precursor_ef": 1.520,
        "scope": "direct_only", "sector": "Demir-Çelik", "markup_type": "standard"
    },
    "7210 49 00 - Kaplanmış / Galvanizli Yassı Çelik": {
        "cn": "7210 49 00", "base_default": 2.510, "eu_benchmark": 1.491, "precursor_ef": 1.610,
        "scope": "direct_only", "sector": "Demir-Çelik", "markup_type": "standard"
    },
    "7214 20 00 - İnşaat Demiri (Nervürlü Çubuk)": {
        "cn": "7214 20 00", "base_default": 2.310, "eu_benchmark": 1.364, "precursor_ef": 1.240,
        "scope": "direct_only", "sector": "Demir-Çelik", "markup_type": "standard"
    },
    "7202 41 - Ferrokrom (Diğer Ülkeler Satırı)": {
        "cn": "7202 41", "base_default": 2.870, "eu_benchmark": 0.900, "precursor_ef": 1.100,
        "scope": "direct_only", "sector": "Demir-Çelik", "markup_type": "standard"
    },
    "7306 30 77 - Kaynaklı Çelik Boru ve İçi Boş Profiller": {
        "cn": "7306 30 77", "base_default": 2.244, "eu_benchmark": 1.491, "precursor_ef": 1.480,
        "scope": "direct_only", "sector": "Demir-Çelik", "markup_type": "standard"
    },
    "7308 90 98 - Çelik Yapılar ve İnşaat Aksamı": {
        "cn": "7308 90 98", "base_default": 5.520, "eu_benchmark": 1.491, "precursor_ef": 2.650,
        "scope": "direct_only", "sector": "Demir-Çelik", "markup_type": "standard"
    },
    "7318 15 88 - Dişli Vida, Cıvata ve Bağlantı Elemanları": {
        "cn": "7318 15 88", "base_default": 5.430, "eu_benchmark": 1.364, "precursor_ef": 2.850,
        "scope": "direct_only", "sector": "Demir-Çelik", "markup_type": "standard"
    },
    "7601 10 90 - İşlenmemiş Alüminyum (Külçe / Biyet)": {
        "cn": "7601 10 90", "base_default": 1.700, "eu_benchmark": 1.423, "precursor_ef": 1.150,
        "scope": "direct_only", "sector": "Alüminyum", "markup_type": "standard"
    },
    "7604 10 90 - Alüminyum Çubuk ve Profiller": {
        "cn": "7604 10 90", "base_default": 3.741, "eu_benchmark": 1.493, "precursor_ef": 2.150,
        "scope": "direct_only", "sector": "Alüminyum", "markup_type": "standard"
    },
    "7606 12 99 - Alüminyum Sac, Levha ve Şeritler": {
        "cn": "7606 12 99", "base_default": 2.714, "eu_benchmark": 1.485, "precursor_ef": 1.780,
        "scope": "direct_only", "sector": "Alüminyum", "markup_type": "standard"
    },
    "7607 11 90 - Alüminyum Folyo": {
        "cn": "7607 11 90", "base_default": 3.071, "eu_benchmark": 1.599, "precursor_ef": 1.950,
        "scope": "direct_only", "sector": "Alüminyum", "markup_type": "standard"
    },
    "2523 10 00 - Gri Çimento Klinkeri (Diğer Ülkeler Satırı)": {
        "cn": "2523 10 00", "base_default": 1.410, "eu_benchmark": 0.666, "precursor_ef": 0.000,
        "scope": "direct_and_indirect", "sector": "Çimento", "markup_type": "standard"
    },
    "2523 29 00 - Gri Portland Çimentosu (Diğer Ülkeler Satırı)": {
        "cn": "2523 29 00", "base_default": 1.440, "eu_benchmark": 0.666, "precursor_ef": 0.000,
        "scope": "direct_and_indirect", "sector": "Çimento", "markup_type": "standard"
    },
    "2523 21 00 - Beyaz Portland Çimentosu (Diğer Ülkeler Satırı)": {
        "cn": "2523 21 00", "base_default": 1.450, "eu_benchmark": 0.859, "precursor_ef": 0.000,
        "scope": "direct_and_indirect", "sector": "Çimento", "markup_type": "standard"
    },
    "2814 10 00 - Susuz Amonyak": {
        "cn": "2814 10 00", "base_default": 2.270, "eu_benchmark": 1.522, "precursor_ef": 0.000,
        "scope": "direct_and_indirect", "sector": "Gübre", "markup_type": "fertiliser"
    },
    "3102 40 10 - Kalsiyum Amonyum Nitrat (CAN Gübresi)": {
        "cn": "3102 40 10", "base_default": 2.230, "eu_benchmark": 0.688, "precursor_ef": 1.150,
        "scope": "direct_and_indirect", "sector": "Gübre", "markup_type": "fertiliser"
    },
    "3102 10 19 - Üre (Azotlu Gübre)": {
        "cn": "3102 10 19", "base_default": 1.510, "eu_benchmark": 0.902, "precursor_ef": 0.850,
        "scope": "direct_and_indirect", "sector": "Gübre", "markup_type": "fertiliser"
    },
    "2804 10 00 - Hidrojen": {
        "cn": "2804 10 00", "base_default": 10.820, "eu_benchmark": 5.089, "precursor_ef": 0.000,
        "scope": "direct_only", "sector": "Hidrojen", "markup_type": "standard"
    },
}

# (EU) 2023/956 Madde 31 Faz Takvimi & Sektör Bazlı Ceza Marjları
YEAR_RULES = {
    "2026 (Kesin Dönem 1. Yıl | Ceza: +%10 [Gübre +%1] | CBAM Faktörü: %97,5)": {
        "std_markup": 1.10, "fert_markup": 1.01, "cbam_factor": 0.975, "year": 2026
    },
    "2027 (Kesin Dönem 2. Yıl | Ceza: +%20 [Gübre +%1] | CBAM Faktörü: %95,0)": {
        "std_markup": 1.20, "fert_markup": 1.01, "cbam_factor": 0.950, "year": 2027
    },
    "2028 (Kesin Dönem 3. Yıl | Ceza: +%30 [Gübre +%1] | CBAM Faktörü: %90,0)": {
        "std_markup": 1.30, "fert_markup": 1.01, "cbam_factor": 0.900, "year": 2028
    },
    "2030 (Orta Vade | Ceza: +%30 [Gübre +%1] | CBAM Faktörü: %51,5)": {
        "std_markup": 1.30, "fert_markup": 1.01, "cbam_factor": 0.515, "year": 2030
    },
    "2034 (Tam Uygulama | Ceza: +%30 [Gübre +%1] | CBAM Faktörü: %0,0)": {
        "std_markup": 1.30, "fert_markup": 1.01, "cbam_factor": 0.000, "year": 2034
    },
}

# --- AKILLI FATURA METİN & GÖRSEL OCR OKUYUCU ---
def parse_energy_text(text):
    kwh_matches = re.findall(r"([\d\.]+(?:,\d+)?)\s*(?:kWh|KWH|kwh)", text)
    m3_matches = re.findall(r"([\d\.]+(?:,\d+)?)\s*(?:m3|m³|Sm3|SM3)", text)

    def parse_tr_float(val_str):
        cleaned = val_str.replace(".", "").replace(",", ".")
        try:
            return float(cleaned)
        except ValueError:
            return 0.0

    kwh_vals = [parse_tr_float(x) for x in kwh_matches if parse_tr_float(x) > 10]
    m3_vals = [parse_tr_float(x) for x in m3_matches if parse_tr_float(x) > 5]

    detected_kwh = max(kwh_vals) if kwh_vals else None
    detected_m3 = max(m3_vals) if m3_vals else None
    return detected_kwh, detected_m3

def extract_energy_from_file(uploaded_file):
    try:
        name_lower = uploaded_file.name.lower()
        if name_lower.endswith(".pdf"):
            reader = PdfReader(uploaded_file)
            full_text = "\n".join([page.extract_text() or "" for page in reader.pages])
            return parse_energy_text(full_text)
        elif OCR_AVAILABLE and any(name_lower.endswith(ext) for ext in [".png", ".jpg", ".jpeg"]):
            ocr = RapidOCR()
            img = Image.open(uploaded_file).convert("RGB")
            result, _ = ocr(np.array(img))
            if result:
                ocr_text = "\n".join([line[1] for line in result])
                return parse_energy_text(ocr_text)
        return None, None
    except Exception:
        return None, None

# --- MADDE 31 & (EU) 2025/2620 UYUMLU HESAPLAMA MOTORU ---
def calculate_cbam_row(
    gtip_label, tonnage, elec_kwh, gas_m3,
    ets_price, local_carbon_price, grid_ef,
    yr_rule, supplier_reduction_pct
):
    info = GTIP_2026_DB[gtip_label]

    # Sektöre göre doğru ceza artışı (Gübre +%1, diğerleri +%10/+%20/+%30)
    markup = yr_rule["fert_markup"] if info["markup_type"] == "fertiliser" else yr_rule["std_markup"]
    markup_pct = int(round((markup - 1.0) * 100))
    cbam_factor = yr_rule["cbam_factor"]

    # Tesis Emisyonları (tCO2e)
    scope1_gas = gas_m3 * NAT_GAS_EF
    scope2_elec = (elec_kwh / 1000.0) * grid_ef
    precursor_em = tonnage * info["precursor_ef"]

    if info["scope"] == "direct_only":
        actual_billable_em = scope1_gas + precursor_em
        scope_note = "Yalnızca Doğrudan (Scope 1 + Precursor) — Elektrik Hariç (Annex II)"
    else:
        actual_billable_em = scope1_gas + scope2_elec + precursor_em
        scope_note = "Doğrudan + Dolaylı (Scope 1 + Scope 2 + Precursor) Dahil"

    # Tedarikçi EAF / Gerçek Veri İyileştirmesi
    best_case_em = max(0.0, actual_billable_em - (precursor_em * (supplier_reduction_pct / 100.0)))

    # Kanıtsız Beyan (2026 TR Baz Default * Yıllık Sektör Marjı)
    default_ef_with_markup = info["base_default"] * markup
    default_total_em = tonnage * default_ef_with_markup

    # Resmi Formül ((EU) 2025/2620): FAA = CBAM_factor * CSCF (1.00) * B_Sütunu_Benchmark * Tonaj
    free_allocation_deduction = tonnage * info["eu_benchmark"] * cbam_factor

    cert_default = max(0.0, default_total_em - free_allocation_deduction)
    cert_actual = max(0.0, actual_billable_em - free_allocation_deduction)
    cert_best = max(0.0, best_case_em - free_allocation_deduction)

    # Net Sertifika Birim Fiyatı (AB CBAM Sertifika Fiyatı - TR-ETS Mahsubu)
    net_carbon_price = max(0.0, ets_price - local_carbon_price)

    tax_default = cert_default * net_carbon_price
    tax_actual = cert_actual * net_carbon_price
    tax_best = cert_best * net_carbon_price

    direct_savings = max(0.0, tax_default - tax_actual)
    supplier_bonus = max(0.0, tax_actual - tax_best)

    return {
        "cn": info["cn"],
        "sector": info["sector"],
        "scope_note": scope_note,
        "markup_pct": markup_pct,
        "cbam_factor": cbam_factor,
        "scope1_gas": scope1_gas,
        "scope2_elec": scope2_elec,
        "precursor_em": precursor_em,
        "eu_benchmark": info["eu_benchmark"],
        "free_alloc_deduction": free_allocation_deduction,
        "base_default": info["base_default"],
        "default_ef_markup": default_ef_with_markup,
        "default_total_em": default_total_em,
        "actual_billable_em": actual_billable_em,
        "best_case_em": best_case_em,
        "cert_default": cert_default,
        "cert_actual": cert_actual,
        "cert_best": cert_best,
        "tax_default": tax_default,
        "tax_actual": tax_actual,
        "tax_best": tax_best,
        "direct_savings": direct_savings,
        "supplier_bonus": supplier_bonus,
        "net_carbon_price": net_carbon_price,
        "is_de_minimis_exempt": tonnage < 50.0,
    }

# --- AB CBAM REGISTRY XML ÜRETİCİSİ ---
def generate_cbam_xml(company_name, broker_name, year, results_list):
    root = ET.Element("CBAMDefinitiveDeclaration", xmlns="urn:eu:taxud:cbam:definitive:v2026")
    header = ET.SubElement(root, "DeclarationHeader")
    ET.SubElement(header, "ReportingYear").text = str(year)
    ET.SubElement(header, "GeneratedAtUTC").text = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    ET.SubElement(header, "InstallationOperatorName").text = company_name
    ET.SubElement(header, "CountryOfOrigin").text = "TR"
    ET.SubElement(header, "AuthorisedDeclarantRepresentative").text = broker_name
    ET.SubElement(header, "RegulatoryReference").text = "Reg (EU) 2023/956 Art.2a, Art.7, Art.31; IR (EU) 2025/2620; IR (EU) 2025/2621; IR (EU) 2026/1740"

    goods_el = ET.SubElement(root, "CBAMImportedGoods")
    for idx, item in enumerate(results_list, start=1):
        good = ET.SubElement(goods_el, "GoodItem", sequence=str(idx))
        ET.SubElement(good, "CNCode").text = str(item["cn"])
        ET.SubElement(good, "Sector").text = str(item["sector"])
        ET.SubElement(good, "NetMassTonnes").text = f"{item['tonnage']:.3f}"
        ET.SubElement(good, "ActualEmbeddedEmissions_tCO2e").text = f"{item['actual_billable_em']:.4f}"
        ET.SubElement(good, "ColumnBBenchmark_tCO2e_per_t").text = f"{item['eu_benchmark']:.3f}"
        ET.SubElement(good, "FreeAllocationAdjustment_tCO2e").text = f"{item['free_alloc_deduction']:.4f}"
        ET.SubElement(good, "NetCBAMCertificatesToSurrender").text = f"{item['cert_actual']:.4f}"
        ET.SubElement(good, "NetCertificateLiability_EUR").text = f"{item['tax_actual']:.2f}"
        ET.SubElement(good, "AvoidedDefaultPenalty_EUR").text = f"{item['direct_savings']:.2f}"

    rough_string = ET.tostring(root, "utf-8")
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ").encode("utf-8")

# --- BEYAZ ETİKET (WHITE-LABEL) KURUMSAL PDF ÜRETİCİ ---
def setup_unicode_font(pdf):
    if os.path.exists("DejaVuSans.ttf"):
        pdf.add_font("TrFont", "", "DejaVuSans.ttf")
        pdf.add_font("TrFont", "B", "DejaVuSans-Bold.ttf" if os.path.exists("DejaVuSans-Bold.ttf") else "DejaVuSans.ttf")
        pdf.add_font("TrFont", "I", "DejaVuSans-Oblique.ttf" if os.path.exists("DejaVuSans-Oblique.ttf") else "DejaVuSans.ttf")
        return "TrFont"
    return "Helvetica"

def generate_corporate_pdf(meta, res, logo_bytes=None):
    pdf = FPDF()
    pdf.add_page()
    font_name = setup_unicode_font(pdf)

    temp_logo_path = None
    if logo_bytes:
        try:
            img = Image.open(io.BytesIO(logo_bytes))
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                img.save(tmp.name, format="PNG")
                temp_logo_path = tmp.name
            pdf.image(temp_logo_path, x=12, y=10, w=25)
            pdf.set_xy(40, 12)
        except Exception:
            pdf.set_xy(10, 12)
    else:
        pdf.set_xy(10, 12)

    pdf.set_fill_color(15, 23, 42)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(font_name, "B", 11.2)
    pdf.cell(0, 11, "TR-CBAM KESİN DÖNEM PRE-AUDIT & MADDE 31 SERTİFİKA RAPORU", ln=True, align="C", fill=True)

    pdf.set_text_color(80, 80, 80)
    pdf.set_font(font_name, "", 8)
    pdf.cell(0, 6, f"Düzenlenme: {datetime.now().strftime('%d.%m.%Y %H:%M')}  |  Belge No: SKDM-{int(datetime.now().timestamp())}  |  Beyan Yılı: {meta['year']}", ln=True, align="C")
    pdf.ln(3.5)

    # 1. Sevkiyat & Kapsam
    pdf.set_text_color(0, 0, 0)
    pdf.set_font(font_name, "B", 9.6)
    pdf.cell(0, 6.2, "1. İHRACATÇI KURULUŞ VE MEVZUAT KAPSAM BİLGİLERİ", ln=True)
    pdf.set_font(font_name, "", 8.6)
    pdf.cell(0, 5.0, f"İhracatçı Kuruluş: {meta['company_name']}", ln=True)
    pdf.cell(0, 5.0, f"Ürün / CN (GTİP) Kodu: {meta['gtip_label']}  ({res['sector']} Sektörü)", ln=True)
    pdf.cell(0, 5.0, f"AB Sevkiyat Miktarı: {meta['tonnage']:,.2f} Ton  |  Net Sertifika Fiyatı: {res['net_carbon_price']:.2f} EUR/tCO2e", ln=True)
    pdf.cell(0, 5.0, f"Emisyon Kapsam Kuralı: {res['scope_note']}", ln=True)
    pdf.cell(0, 5.0, f"Enerji Kanıt Belgesi: {meta['invoice_status']}", ln=True)
    pdf.ln(3)

    # 2. Finansal Kıyaslama Tablosu
    pdf.set_fill_color(241, 245, 249)
    pdf.set_font(font_name, "B", 9.6)
    pdf.cell(0, 7.2, "2. AB SKDM SERTİFİKA YÜKÜMLÜLÜĞÜ KIYASLAMASI ((EU) 2025/2620 FORMÜLÜ)", ln=True, fill=True)
    pdf.ln(1.5)

    pdf.set_font(font_name, "B", 8.0)
    pdf.cell(78, 6.8, "Beyan Senaryosu", border=1)
    pdf.cell(36, 6.8, "Gömülü Emisyon", border=1, align="C")
    pdf.cell(36, 6.8, "Net Sertifika (Adet)", border=1, align="C")
    pdf.cell(40, 6.8, "Ödenecek Tutar", border=1, align="R", ln=True)

    pdf.set_font(font_name, "", 8.0)
    pdf.cell(78, 6.8, f"A: Kanıtsız Beyan (TR Default +%{res['markup_pct']} Marj)", border=1)
    pdf.cell(36, 6.8, f"{res['default_total_em']:,.2f} t", border=1, align="C")
    pdf.cell(36, 6.8, f"{res['cert_default']:,.2f}", border=1, align="C")
    pdf.cell(40, 6.8, f"{res['tax_default']:,.2f} EUR", border=1, align="R", ln=True)

    pdf.cell(78, 6.8, "B: Kanıtlı Tesis Verisi (Fatura & Sayaç)", border=1)
    pdf.cell(36, 6.8, f"{res['actual_billable_em']:,.2f} t", border=1, align="C")
    pdf.cell(36, 6.8, f"{res['cert_actual']:,.2f}", border=1, align="C")
    pdf.cell(40, 6.8, f"{res['tax_actual']:,.2f} EUR", border=1, align="R", ln=True)

    pdf.cell(78, 6.8, "C: Tesis + Tedarikçi Doğrulanmış Veri", border=1)
    pdf.cell(36, 6.8, f"{res['best_case_em']:,.2f} t", border=1, align="C")
    pdf.cell(36, 6.8, f"{res['cert_best']:,.2f}", border=1, align="C")
    pdf.cell(40, 6.8, f"{res['tax_best']:,.2f} EUR", border=1, align="R", ln=True)

    pdf.set_font(font_name, "B", 9.0)
    pdf.set_fill_color(220, 252, 231)
    pdf.cell(150, 8.2, "DOĞRUDAN KORUNAN NET SKDM SERTİFİKA TASARRUFU (A - B)", border=1, fill=True)
    pdf.cell(40, 8.2, f"+{res['direct_savings']:,.2f} EUR", border=1, align="R", ln=True, fill=True)
    pdf.ln(3)

    # 3. Emisyon & Ücretsiz Tahsisat Düşüm Detayı
    pdf.set_font(font_name, "B", 9.6)
    pdf.cell(0, 6.2, "3. TESİS EMİSYON KIRILIMI VE MADDE 31 ÜCRETSİZ TAHSİSAT DÜŞÜMÜ", ln=True)
    pdf.set_font(font_name, "", 8.4)
    pdf.cell(0, 4.8, f"• Scope 1 (Doğalgaz Yanma Emisyonu - 56,1 tCO2/TJ): {res['scope1_gas']:,.2f} tCO2e", ln=True)
    pdf.cell(0, 4.8, f"• Scope 2 (Elektrik Tüketimi - ETKB Katsayısı): {res['scope2_elec']:,.2f} tCO2e ({'Kapsam Dışı' if 'Hariç' in res['scope_note'] else 'Hesaba Dahil'})", ln=True)
    pdf.cell(0, 4.8, f"• Precursor (Öncül Hammadde Gömülü Emisyonu): {res['precursor_em']:,.2f} tCO2e", ln=True)
    pdf.cell(0, 4.8, f"• AB Ücretsiz Tahsisat Düşümü (CBAM Faktörü %{res['cbam_factor']*100:.1f} x Benchmark {res['eu_benchmark']:.3f} t/t): -{res['free_alloc_deduction']:,.2f} tCO2e", ln=True)
    pdf.ln(2)

    if res["supplier_bonus"] > 0:
        pdf.set_fill_color(254, 243, 199)
        pdf.set_font(font_name, "I", 7.8)
        pdf.multi_cell(0, 4.5, f"TEDARİKÇİ OPTİMİZASYON NOTU: Öncül hammadde tedarikçinizden (EU) 2025/2547 uyarınca akredite doğrulanmış emisyon beyanı alınması halinde, bu sevkiyatta ilave +{res['supplier_bonus']:,.2f} EUR daha net sertifika maliyeti avantajı elde edilebilir.", border=1, fill=True)
        pdf.ln(3)

    # 4. Kaşe & Resmi Mevzuat Dipnotu
    pdf.set_font(font_name, "B", 8.5)
    pdf.cell(0, 5.2, f"Raporu Düzenleyen Yetkili Kurum / Gümrük Müşavirliği: {meta['broker_name']}", ln=True)
    pdf.ln(1.2)
    pdf.set_font(font_name, "", 7.1)
    pdf.set_text_color(75, 75, 75)
    pdf.multi_cell(0, 3.7, "Resmi Mevzuat ve Katsayı Dayanağı: Hesaplamalarda (EU) 2023/956 sayılı AB SKDM Tüzüğü (Madde 2a, Annex II, Annex IV ve Madde 31), ücretsiz tahsisat kıyas değerlerini belirleyen (EU) 2025/2620 sayılı Uygulama Tüzüğü, Türkiye özel kesin dönem varsayılan değerlerini düzenleyen (EU) 2025/2621 ve (EU) 2026/1740 sayılı Tüzükler, yakıt faktörlerini düzenleyen (EU) 2023/1773 sayılı Tüzük ile T.C. ETKB Ulusal Elektrik Şebekesi Emisyon Faktörleri esas alınmıştır. Kesin gümrük beyanında akredite doğrulayıcı ((EU) 2025/2547) onayı zorunludur.")
    pdf.ln(1.0)
    pdf.multi_cell(0, 3.7, "Zero-Retention Veri Güvenliği Beyanı: Yüklenen fatura, enerji tüketim ve sevkiyat verileri hiçbir veritabanında kalıcı olarak depolanmaz; oturum belleğinde anlık işlenir ve oturum kapandığında kalıcı olarak silinir.")

    out_bytes = bytes(pdf.output())
    if temp_logo_path and os.path.exists(temp_logo_path):
        try:
            os.remove(temp_logo_path)
        except Exception:
            pass
    return out_bytes

# --- YAN MENÜ (GLOBAL PARAMETRELER) ---
with st.sidebar:
    st.header("⚙️ Kurumsal & Mevzuat Ayarları")
    broker_name = st.text_input("Gümrük Müşavirliği / Danışmanlık Unvanı", value="İÜ Girişim Hub — R&D Division")
    broker_logo = st.file_uploader("Kurumsal Logo Yükle (PDF Anteti İçin)", type=["png", "jpg", "jpeg"])

    st.divider()
    st.subheader("📅 AB Madde 31 & Piyasa Parametreleri")
    year_choice = st.selectbox("Beyan Yılı & CBAM Faktörü", list(YEAR_RULES.keys()))
    grid_choice = st.selectbox("Elektrik Emisyon Faktörü Kaynağı", list(ETKB_GRID_FACTORS.keys()))
    ets_price = st.number_input(
        "AB CBAM Sertifika Fiyatı (€ / tCO₂e)",
        min_value=10.0, value=82.32, step=0.5,
        help="Avrupa Komisyonu 2026 Q3 Resmi CBAM Sertifika Fiyatı: 82,32 € / tCO2e"
    )
    local_carbon_price = st.number_input(
        "TR-ETS / Türkiye İçi Ödenen Karbon Bedeli (€ / tCO₂e)",
        min_value=0.0, value=0.0, step=1.0,
        help="Türkiye Emisyon Ticaret Sistemi (TR-ETS) kapsamında ödenen karbon fiyatı AB SKDM sertifika yükünden birebir mahsup edilir."
    )
    supplier_reduction_pct = st.slider(
        "Tedarikçi EAF / Hurda Bazlı Emisyon İyileşme Oranı (%)",
        min_value=0, max_value=70, value=40, step=5,
        help="Öncül hammadde tedarikçisinin elektrik ark ocağı (EAF) veya düşük karbonlu üretim verisi sunması halindeki tahmini emisyon düşüşü."
    )

    st.divider()
    st.caption("🔒 **Zero-Retention:** Tüm veriler yalnızca aktif oturum belleğinde işlenir; veritabanına kaydedilmez.")

# --- ANA EKRAN BAŞLIĞI ---
st.title("TR-CBAM v3.0 | Resmi Mevzuat SKDM Uyum & Denetim Kokpiti")
st.markdown("**(EU) 2025/2620 Resmi Benchmarklar + (EU) 2025/2621 & 2026/1740 Türkiye Katsayıları (20 Ürün Grubu) + Madde 2a Muafiyet Kontrolü**")

tab1, tab2, tab3 = st.tabs([
    "📄 Tekil Sevkiyat & Fatura OCR Analizi",
    "📊 Çoklu Kalem (Excel) & AB XML Çıktısı",
    "📚 2026 Türkiye Resmi Katsayı & Benchmark Kataloğu (20 Kalem)"
])

# ==============================================================================
# SEKME 1: TEKİL SEVKİYAT + FATURA OCR + GRAFİKLER + KURUMSAL PDF + XML
# ==============================================================================
with tab1:
    col_left, col_right = st.columns([1.15, 1])

    with col_left:
        st.subheader("1. Tesis ve Fatura Verileri")
        invoice_file = st.file_uploader(
            "Opsiyonel: Elektrik veya Doğalgaz Faturası Yükle (PDF / JPG / PNG - Otomatik Okuma)",
            type=["pdf", "png", "jpg", "jpeg"]
        )

        auto_kwh, auto_m3 = 45000.0, 3200.0
        invoice_status = "Kanıt Belgesi Henüz Yüklenmedi (Ön Simülasyon)"
        if invoice_file is not None:
            det_kwh, det_m3 = extract_energy_from_file(invoice_file)
            invoice_status = f"Doğrulandı: {invoice_file.name}"
            if det_kwh:
                auto_kwh = det_kwh
                st.success(f"Faturadan otomatik elektrik tüketimi okundu: **{det_kwh:,.2f} kWh**")
            if det_m3:
                auto_m3 = det_m3
                st.success(f"Faturadan otomatik doğalgaz tüketimi okundu: **{det_m3:,.2f} m³**")
            if not det_kwh and not det_m3:
                st.info("Fatura dosyası eklendi; tüketim değerlerini aşağıdan kontrol edip onaylayabilirsiniz.")

        with st.form("single_cbam_form"):
            company_name = st.text_input("İhracatçı Firma Unvanı", value="Örnek Metal ve Çelik Sanayi A.Ş.")
            gtip_label = st.selectbox("Ürün CN (GTİP) Kodu (20 Resmi Kategori)", list(GTIP_2026_DB.keys()), index=1)
            c1, c2, c3 = st.columns(3)
            tonnage = c1.number_input("AB Sevkiyatı (Ton)", min_value=1.0, value=120.0, step=10.0)
            elec_kwh = c2.number_input("Elektrik (kWh)", min_value=0.0, value=float(auto_kwh), step=500.0)
            gas_m3 = c3.number_input("Doğalgaz (m³)", min_value=0.0, value=float(auto_m3), step=100.0)

            run_btn = st.form_submit_button("Kesin Dönem Analizini & Belgeleri Üret", use_container_width=True, type="primary")

    if run_btn:
        yr = YEAR_RULES[year_choice]
        grid_ef = ETKB_GRID_FACTORS[grid_choice]
        res = calculate_cbam_row(
            gtip_label, tonnage, elec_kwh, gas_m3,
            ets_price, local_carbon_price, grid_ef,
            yr, supplier_reduction_pct
        )

        with col_right:
            st.subheader("2. Finansal Etki (Madde 31 Net Sertifika Yükü)")
            if res["is_de_minimis_exempt"]:
                st.warning("⚠️ **Madde 2a De Minimis Uyarısı:** AB ithalatçısının yıllık toplam CBAM mal ithalatı **50 tonun altındaysa** gümrük beyanında muafiyet işaretlenir ve sertifika teslimi yapılmaz.")

            m1, m2, m3 = st.columns(3)
            m1.metric("Kanıtsız Ceza (A)", f"{res['tax_default']:,.0f} €", help=f"Sertifika: {res['cert_default']:,.2f} adet")
            m2.metric("Kanıtlı Beyan (B)", f"{res['tax_actual']:,.0f} €", help=f"Sertifika: {res['cert_actual']:,.2f} adet")
            m3.metric("Net Vergi Tasarrufu", f"+{res['direct_savings']:,.0f} €", delta=f"Ek Tedarikçi Fırsatı: +{res['supplier_bonus']:,.0f} €")

            fig = go.Figure(data=[
                go.Bar(
                    name="Net Ödenecek Tutar (€)",
                    x=["A: Kanıtsız Default", "B: Tesis Gerçek Verisi", "C: Tedarikçi Verisi Dahil"],
                    y=[res["tax_default"], res["tax_actual"], res["tax_best"]],
                    marker_color=["#ef4444", "#3b82f6", "#10b981"],
                    text=[f"{res['tax_default']:,.0f} €", f"{res['tax_actual']:,.0f} €", f"{res['tax_best']:,.0f} €"],
                    textposition="auto"
                )
            ])
            fig.update_layout(
                title=f"{yr['year']} Yılı Net AB SKDM Sertifika Maliyeti Kıyaslaması (€)",
                height=310,
                margin=dict(l=20, r=20, t=40, b=20),
                yaxis_title="Euro (€)"
            )
            st.plotly_chart(fig, use_container_width=True)

        st.divider()
        st.info(
            f"**Mevzuat Kapsamı:** {res['scope_note']} | "
            f"**TR Ceza Katsayısı (+%{res['markup_pct']}):** `{res['default_ef_markup']:.4f} tCO₂e/t` | "
            f"**(EU) 2025/2620 Benchmark:** `{res['eu_benchmark']:.3f} t/t` | "
            f"**Ücretsiz Tahsisat Düşümü:** `-{res['free_alloc_deduction']:,.2f} tCO₂e`"
        )

        b_col1, b_col2 = st.columns(2)
        meta_data = {
            "company_name": company_name,
            "gtip_label": gtip_label,
            "tonnage": tonnage,
            "broker_name": broker_name,
            "invoice_status": invoice_status,
            "year": yr["year"]
        }
        logo_bytes = broker_logo.getvalue() if broker_logo else None
        pdf_bytes = generate_corporate_pdf(meta_data, res, logo_bytes)

        xml_item = dict(res)
        xml_item["tonnage"] = tonnage
        xml_bytes = generate_cbam_xml(company_name, broker_name, yr["year"], [xml_item])

        b_col1.download_button(
            label="📥 Kurumsal Pre-Audit PDF Raporunu İndir",
            data=pdf_bytes,
            file_name=f"TR_CBAM_{yr['year']}_Rapor_{company_name.replace(' ', '_')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
        b_col2.download_button(
            label="🧾 AB CBAM Portalı Resmi XML Dosyasını İndir",
            data=xml_bytes,
            file_name=f"EU_CBAM_Registry_{res['cn'].replace(' ', '')}_{yr['year']}.xml",
            mime="application/xml",
            use_container_width=True
        )

# ==============================================================================
# SEKME 2: ÇOKLU KALEM (EXCEL / CSV) TOPLU HESAPLAMA & XML
# ==============================================================================
with tab2:
    st.subheader("Çoklu İhracat Kalemi Toplu Analiz ve AB XML Üretimi")
    st.markdown("Onlarca kalemlik ihracat faturasını tek seferde hesaplamak için aşağıdaki örnek şablonu indirip doldurun ve yükleyin.")

    sample_df = pd.DataFrame([
        {"Urun_Adi": "Sıcak Haddelenmiş Sac", "GTIP_Kodu": "7208 51 20 - Sıcak Haddelenmiş Yassı Çelik Ürünler", "Tonaj": 120.0, "Elektrik_kWh": 45000.0, "Dogalgaz_m3": 3200.0},
        {"Urun_Adi": "Çelik Bağlantı Cıvatası", "GTIP_Kodu": "7318 15 88 - Dişli Vida, Cıvata ve Bağlantı Elemanları", "Tonaj": 35.0, "Elektrik_kWh": 18000.0, "Dogalgaz_m3": 1100.0},
        {"Urun_Adi": "Alüminyum Profil", "GTIP_Kodu": "7604 10 90 - Alüminyum Çubuk ve Profiller", "Tonaj": 50.0, "Elektrik_kWh": 28000.0, "Dogalgaz_m3": 1900.0},
    ])

    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        sample_df.to_excel(writer, index=False, sheet_name="Sevkiyat_Listesi")

    st.download_button(
        label="📄 Örnek Çoklu Sevkiyat Excel Şablonunu İndir (.xlsx)",
        data=excel_buf.getvalue(),
        file_name="TR_CBAM_Toplu_Yukleme_Sablonu.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    uploaded_excel = st.file_uploader("Doldurulmuş Excel Dosyasını Yükle (.xlsx)", type=["xlsx"])
    if uploaded_excel:
        try:
            df_in = pd.read_excel(uploaded_excel)
            yr = YEAR_RULES[year_choice]
            grid_ef = ETKB_GRID_FACTORS[grid_choice]
            batch_results = []
            summary_rows = []

            for _, row in df_in.iterrows():
                g_label = row["GTIP_Kodu"] if row["GTIP_Kodu"] in GTIP_2026_DB else list(GTIP_2026_DB.keys())[1]
                r = calculate_cbam_row(
                    g_label, float(row["Tonaj"]), float(row["Elektrik_kWh"]), float(row["Dogalgaz_m3"]),
                    ets_price, local_carbon_price, grid_ef, yr, supplier_reduction_pct
                )
                r["tonnage"] = float(row["Tonaj"])
                batch_results.append(r)
                summary_rows.append({
                    "Ürün": row.get("Urun_Adi", g_label),
                    "CN Kodu": r["cn"],
                    "Tonaj": float(row["Tonaj"]),
                    "Kanıtsız Ceza (€)": round(r["tax_default"], 2),
                    "Kanıtlı Vergi (€)": round(r["tax_actual"], 2),
                    "Net Tasarruf (€)": round(r["direct_savings"], 2),
                })

            res_df = pd.DataFrame(summary_rows)
            st.dataframe(res_df, use_container_width=True)

            tot_sav = res_df["Net Tasarruf (€)"].sum()
            st.success(f"**Toplam Sevkiyat Tasarrufu ({yr['year']}):** `{tot_sav:,.2f} €`")

            batch_xml = generate_cbam_xml("Toplu İhracat Beyanı", broker_name, yr["year"], batch_results)
            st.download_button(
                label="🧾 Tüm Kalemler İçin Tek AB CBAM XML Dosyasını İndir",
                data=batch_xml,
                file_name=f"EU_CBAM_Batch_Registry_{yr['year']}.xml",
                mime="application/xml",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"Excel dosyası işlenirken hata oluştu. Lütfen şablon sütun adlarını koruyun. Detay: {e}")

# ==============================================================================
# SEKME 3: 2026 TÜRKİYE RESMİ KATSAYI & BENCHMARK REFERANS TABLOSU (20 KALEM)
# ==============================================================================
with tab3:
    st.subheader("(EU) 2025/2621, (EU) 2026/1740 ve (EU) 2025/2620 Resmi Türkiye Kataloğu")
    ref_list = []
    for k, v in GTIP_2026_DB.items():
        m_2026 = 1.01 if v["markup_type"] == "fertiliser" else 1.10
        ref_list.append({
            "CN / GTİP": v["cn"],
            "Ürün Tanımı": k.split(" - ")[1],
            "Sektör": v["sector"],
            "2026 TR Baz Değer (tCO₂e/t)": v["base_default"],
            "2026 Ceza Marjı": "+%1" if v["markup_type"] == "fertiliser" else "+%10",
            "2026 Marjl› Katsay›": round(v["base_default"] * m_2026, 4),
            "B Sütunu Kıyas Değeri ((EU) 2025/2620)": v["eu_benchmark"],
            "Mali Yükümlülük Kapsamı": "Sadece Doğrudan (Direct)" if v["scope"] == "direct_only" else "Doğrudan + Dolaylı (Elektrik)"
        })
    st.dataframe(pd.DataFrame(ref_list), use_container_width=True)
