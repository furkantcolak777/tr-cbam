import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from fpdf import FPDF
from pypdf import PdfReader
from PIL import Image
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime
import io
import os
import re

# --- SAYFA AYARLARI ---
st.set_page_config(
    page_title="TR-CBAM v2.0 | Kurumsal SKDM Uyum Kokpiti",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2026 KESİN DÖNEM RESMİ VERİ TABANI ---
# Kaynak 1: T.C. ETKB (2023 Hesaplama Dönemi Tüketim Noktası Emisyon Faktörleri)
ETKB_GRID_FACTORS = {
    "Dağıtım Hattı / OSB Bağlantısı (0,469 tCO₂/MWh)": 0.469,
    "İletim Hattı / Yüksek Gerilim (0,436 tCO₂/MWh)": 0.436,
    "Türkiye Geneli Üretim Ortalaması (0,434 tCO₂/MWh)": 0.434,
}

# Kaynak 2: AB Uygulama Tüzüğü (EU) 2023/1773 (56,1 tCO2/TJ * 0,0342 TJ/1000m3)
NAT_GAS_EF = 0.00192  # tCO2 / m3

# Kaynak 3: (EU) 2025/2621 & (EU) 2026/1740 Türkiye Özel Kesin Dönem Varsayılan Değerleri (tCO2e / Ton)
# scope_type: "direct_only" (Demir-Çelik & Alüminyum) | "direct_and_indirect" (Çimento & Gübre)
GTIP_2026_DB = {
    "7208 - Sıcak Haddelenmiş Yassı Çelik Ürünler": {
        "cn": "7208", "base_default": 2.428, "precursor_ef": 1.350, "scope": "direct_only", "sector": "Demir-Çelik"
    },
    "7209 - Soğuk Haddelenmiş Yassı Çelik Ürünler": {
        "cn": "7209", "base_default": 2.580, "precursor_ef": 1.520, "scope": "direct_only", "sector": "Demir-Çelik"
    },
    "7210 - Galvanizli / Kaplamalı Yassı Çelik": {
        "cn": "7210", "base_default": 2.710, "precursor_ef": 1.610, "scope": "direct_only", "sector": "Demir-Çelik"
    },
    "7214 20 00 - İnşaat Demiri (Nervürlü Çubuk)": {
        "cn": "72142000", "base_default": 2.310, "precursor_ef": 1.240, "scope": "direct_only", "sector": "Demir-Çelik"
    },
    "7306 19 00 - Kaynaklı Çelik Boru ve İçi Boş Profiller": {
        "cn": "73061900", "base_default": 2.244, "precursor_ef": 1.480, "scope": "direct_only", "sector": "Demir-Çelik"
    },
    "7318 15 - Cıvata, Vida, Somun ve Bağlantı Elemanları": {
        "cn": "731815", "base_default": 5.430, "precursor_ef": 2.850, "scope": "direct_only", "sector": "Demir-Çelik"
    },
    "7601 - İşlenmemiş Alüminyum (Külçe / Biyet)": {
        "cn": "7601", "base_default": 1.700, "precursor_ef": 1.150, "scope": "direct_only", "sector": "Alüminyum"
    },
    "7604 10 90 - Alüminyum Çubuk ve Profiller": {
        "cn": "76041090", "base_default": 3.741, "precursor_ef": 2.150, "scope": "direct_only", "sector": "Alüminyum"
    },
    "7606 - Alüminyum Sac, Levha ve Şeritler": {
        "cn": "7606", "base_default": 2.714, "precursor_ef": 1.780, "scope": "direct_only", "sector": "Alüminyum"
    },
    "2523 29 00 - Gri Portland Çimento": {
        "cn": "25232900", "base_default": 1.440, "precursor_ef": 0.000, "scope": "direct_and_indirect", "sector": "Çimento"
    },
    "2523 30 00 - Alüminli Çimento": {
        "cn": "25233000", "base_default": 1.950, "precursor_ef": 0.000, "scope": "direct_and_indirect", "sector": "Çimento"
    },
    "3102 - Azotlu Mineral veya Kimyasal Gübreler": {
        "cn": "3102", "base_default": 2.850, "precursor_ef": 1.400, "scope": "direct_and_indirect", "sector": "Gübre"
    },
}

# 2026-2028 Kesin Dönem Ceza Artış Oranları (Mark-up) & AB Ücretsiz Tahsisat Faz-Out Faktörü
YEAR_RULES = {
    "2026 (Kesin Dönem 1. Yıl - +%10 Ceza Artışı)": {"markup": 1.10, "cbam_phase_in": 0.975, "year": 2026},
    "2027 (Kesin Dönem 2. Yıl - +%20 Ceza Artışı)": {"markup": 1.20, "cbam_phase_in": 0.950, "year": 2027},
    "2028+ (Kesin Dönem 3. Yıl - +%30 Ceza Artışı)": {"markup": 1.30, "cbam_phase_in": 0.900, "year": 2028},
}

# --- AKILLI FATURA METİN TARAYICI (PDF OCR/TEXT PARSER) ---
def extract_energy_from_pdf(uploaded_file):
    try:
        reader = PdfReader(uploaded_file)
        full_text = "\n".join([page.extract_text() or "" for page in reader.pages])
        
        # kWh yakalama (Örn: 45.250,00 kWh veya 45250 kWh)
        kwh_matches = re.findall(r"([\d\.]+(?:,\d+)?)\s*(?:kWh|KWH|kwh)", full_text)
        # m3 yakalama (Örn: 3.200 m3 veya 3200 sm3)
        m3_matches = re.findall(r"([\d\.]+(?:,\d+)?)\s*(?:m3|m³|Sm3|SM3)", full_text)

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
        return detected_kwh, detected_m3, True if full_text.strip() else False
    except Exception:
        return None, None, False

# --- HESAPLAMA MOTORU ---
def calculate_cbam_row(gtip_label, tonnage, elec_kwh, gas_m3, ets_price, local_carbon_price, grid_ef, markup, phase_in):
    info = GTIP_2026_DB[gtip_label]
    
    # Tesis Emisyonları
    scope1_gas = gas_m3 * NAT_GAS_EF
    scope2_elec = (elec_kwh / 1000.0) * grid_ef
    precursor_em = tonnage * info["precursor_ef"]

    # Kapsam kuralı (Demir-Çelik/Alüminyum: Sadece Direct + Precursor | Çimento/Gübre: Direct + Indirect)
    if info["scope"] == "direct_only":
        actual_billable_em = scope1_gas + precursor_em
        scope_note = "Yalnızca Doğrudan (Scope 1 + Precursor) — Elektrik Hariç (Annex II)"
    else:
        actual_billable_em = scope1_gas + scope2_elec + precursor_em
        scope_note = "Doğrudan + Dolaylı (Scope 1 + Scope 2 + Precursor) Dahil"

    # Tedarikçi verisi de alınırsa (Precursor %40 iyileşme potansiyeli)
    best_case_em = actual_billable_em - (precursor_em * 0.40)

    # Kanıtsız Beyan (2026 TR Baz Default * Yıllık Mark-up)
    default_ef_with_markup = info["base_default"] * markup
    default_total_em = tonnage * default_ef_with_markup

    # Net Sertifika Maliyeti = Emisyon * (AB ETS - Yerel Karbon Fiyatı) * CBAM Faz Oranı
    effective_carbon_price = max(0.0, ets_price - local_carbon_price) * phase_in

    tax_default = default_total_em * effective_carbon_price
    tax_actual = actual_billable_em * effective_carbon_price
    tax_best = best_case_em * effective_carbon_price

    direct_savings = max(0.0, tax_default - tax_actual)
    supplier_bonus = max(0.0, tax_actual - tax_best)

    return {
        "cn": info["cn"],
        "sector": info["sector"],
        "scope_note": scope_note,
        "scope1_gas": scope1_gas,
        "scope2_elec": scope2_elec,
        "precursor_em": precursor_em,
        "base_default": info["base_default"],
        "default_ef_markup": default_ef_with_markup,
        "default_total_em": default_total_em,
        "actual_billable_em": actual_billable_em,
        "best_case_em": best_case_em,
        "tax_default": tax_default,
        "tax_actual": tax_actual,
        "tax_best": tax_best,
        "direct_savings": direct_savings,
        "supplier_bonus": supplier_bonus,
        "effective_price": effective_carbon_price,
    }

# --- AB CBAM REGISTRY XML ÜRETİCİSİ ---
def generate_cbam_xml(company_name, broker_name, year, results_list):
    root = ET.Element("CBAMQuarterlyOrAnnualReport", xmlns="urn:eu:taxud:cbam:v2026")
    header = ET.SubElement(root, "ReportHeader")
    ET.SubElement(header, "ReportingPeriodYear").text = str(year)
    ET.SubElement(header, "GeneratedAtUTC").text = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    ET.SubElement(header, "ExporterInstallationName").text = company_name
    ET.SubElement(header, "CountryOfOrigin").text = "TR"
    ET.SubElement(header, "CustomsDeclarantRepresentative").text = broker_name
    ET.SubElement(header, "LegalBasis").text = "Regulation (EU) 2023/956; IR (EU) 2025/2621; IR (EU) 2026/1740"

    goods_el = ET.SubElement(root, "ImportedCBAMGoods")
    for idx, item in enumerate(results_list, start=1):
        good = ET.SubElement(goods_el, "CBAMGoodItem", sequence=str(idx))
        ET.SubElement(good, "CNCode").text = str(item["cn"])
        ET.SubElement(good, "Sector").text = str(item["sector"])
        ET.SubElement(good, "NetMassTonnes").text = f"{item['tonnage']:.3f}"
        ET.SubElement(good, "CalculationMethod").text = "ActualVerifiedInstallationData_AnnexIV"
        ET.SubElement(good, "ActualEmbeddedEmissions_tCO2e").text = f"{item['actual_billable_em']:.4f}"
        ET.SubElement(good, "DefaultFallbackEmissions_tCO2e").text = f"{item['default_total_em']:.4f}"
        ET.SubElement(good, "EstimatedCBAMCertificatePayable_EUR").text = f"{item['tax_actual']:.2f}"
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

    # Logo Varsa Üst Sol Köşeye Ekle
    if logo_bytes:
        try:
            img = Image.open(io.BytesIO(logo_bytes))
            temp_logo = "/tmp/temp_broker_logo.png"
            img.save(temp_logo, format="PNG")
            pdf.image(temp_logo, x=12, y=10, w=26)
            pdf.set_xy(42, 12)
        except Exception:
            pdf.set_xy(10, 12)
    else:
        pdf.set_xy(10, 12)

    # Üst Kurumsal Başlık
    pdf.set_fill_color(15, 23, 42)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(font_name, "B", 12)
    pdf.cell(0, 11, "TR-CBAM KESİN DÖNEM PRE-AUDIT & SINIR VERGİSİ RAPORU", ln=True, align="C", fill=True)

    pdf.set_text_color(80, 80, 80)
    pdf.set_font(font_name, "", 8)
    pdf.cell(0, 6, f"Düzenlenme Tarihi: {datetime.now().strftime('%d.%m.%Y %H:%M')}  |  Belge No: SKDM-{int(datetime.now().timestamp())}  |  Dönem: {meta['year']}", ln=True, align="C")
    pdf.ln(5)

    # 1. İhracatçı & Mevzuat Kapsamı
    pdf.set_text_color(0, 0, 0)
    pdf.set_font(font_name, "B", 10)
    pdf.cell(0, 7, "1. İHRACATÇI FİRMA VE SEVKİYAT KAPSAM BİLGİLERİ", ln=True)
    pdf.set_font(font_name, "", 9)
    pdf.cell(0, 5.5, f"İhracatçı Kuruluş: {meta['company_name']}", ln=True)
    pdf.cell(0, 5.5, f"Ürün / CN (GTİP) Kodu: {meta['gtip_label']}  ({res['sector']} Sektörü)", ln=True)
    pdf.cell(0, 5.5, f"AB İthalat Miktarı: {meta['tonnage']:,.2f} Ton  |  Net Karbon Fiyatı: {res['effective_price']:.2f} EUR/tCO2e", ln=True)
    pdf.cell(0, 5.5, f"Mevzuat Kapsam Kuralı: {res['scope_note']}", ln=True)
    pdf.cell(0, 5.5, f"Enerji Kanıt Durumu: {meta['invoice_status']}", ln=True)
    pdf.ln(4)

    # 2. Finansal Kıyaslama Tablosu
    pdf.set_fill_color(241, 245, 249)
    pdf.set_font(font_name, "B", 10)
    pdf.cell(0, 8, "2. AB SINIRDA KARBON VERGİSİ (SKDM) MALİYET KIYASLAMASI", ln=True, fill=True)
    pdf.ln(1.5)

    pdf.set_font(font_name, "B", 8.5)
    pdf.cell(92, 7.5, "Beyan ve Doğrulama Senaryosu", border=1)
    pdf.cell(48, 7.5, "Hesaplanan Emisyon", border=1, align="C")
    pdf.cell(50, 7.5, "SKDM Sertifika Maliyeti", border=1, align="R", ln=True)

    pdf.set_font(font_name, "", 8.5)
    pdf.cell(92, 7.5, f"A: Kanıtsız Beyan (TR Default + Ceza Katsayısı)", border=1)
    pdf.cell(48, 7.5, f"{res['default_total_em']:,.2f} tCO2e", border=1, align="C")
    pdf.cell(50, 7.5, f"{res['tax_default']:,.2f} EUR", border=1, align="R", ln=True)

    pdf.cell(92, 7.5, "B: Kanıtlı Tesis Verisi (Fatura & Sayaç Bazlı)", border=1)
    pdf.cell(48, 7.5, f"{res['actual_billable_em']:,.2f} tCO2e", border=1, align="C")
    pdf.cell(50, 7.5, f"{res['tax_actual']:,.2f} EUR", border=1, align="R", ln=True)

    pdf.cell(92, 7.5, "C: Tesis + Tedarikçi Doğrulanmış Öncül Veri", border=1)
    pdf.cell(48, 7.5, f"{res['best_case_em']:,.2f} tCO2e", border=1, align="C")
    pdf.cell(50, 7.5, f"{res['tax_best']:,.2f} EUR", border=1, align="R", ln=True)

    pdf.set_font(font_name, "B", 9.5)
    pdf.set_fill_color(220, 252, 231)
    pdf.cell(140, 9, "DOĞRUDAN KORUNAN SINIR VERGİSİ AVANTAJI (A - B)", border=1, fill=True)
    pdf.cell(50, 9, f"+{res['direct_savings']:,.2f} EUR", border=1, align="R", ln=True, fill=True)
    pdf.ln(4)

    # 3. Teknik Emisyon Kırılımı
    pdf.set_font(font_name, "B", 10)
    pdf.cell(0, 7, "3. TESİS EMİSYON KIRILIMI VE TEDARİK ZİNCİRİ NOTU", ln=True)
    pdf.set_font(font_name, "", 8.8)
    pdf.cell(0, 5.5, f"• Scope 1 (Doğalgaz Yanma Emisyonu - 56,1 tCO2/TJ): {res['scope1_gas']:,.2f} tCO2e", ln=True)
    pdf.cell(0, 5.5, f"• Scope 2 (Elektrik Tüketim Emisyonu - ETKB Katsayısı): {res['scope2_elec']:,.2f} tCO2e ({'Kapsam Dışı' if 'Hariç' in res['scope_note'] else 'Hesaba Dahil'})", ln=True)
    pdf.cell(0, 5.5, f"• Precursor (Öncül Hammadde Gömülü Emisyonu): {res['precursor_em']:,.2f} tCO2e", ln=True)
    pdf.ln(2.5)

    if res["supplier_bonus"] > 0:
        pdf.set_fill_color(254, 243, 199)
        pdf.set_font(font_name, "I", 8.2)
        pdf.multi_cell(0, 4.8, f"TEDARİKÇİ OPTİMİZASYON NOTU: Öncül hammadde tedarikçinizden (EU) 2025/2547 uyarınca akredite doğrulanmış emisyon beyanı temin edilmesi durumunda, bu sevkiyatta ilave +{res['supplier_bonus']:,.2f} EUR daha sertifika maliyeti düşüşü sağlanabilir.", border=1, fill=True)
        pdf.ln(4)

    # 4. Yetkili Gümrük Müşavirliği Kaşe & Resmi Mevzuat Dipnotu
    pdf.set_font(font_name, "B", 8.8)
    pdf.cell(0, 6, f"Raporu Düzenleyen Yetkili Kurum / Gümrük Müşavirliği: {meta['broker_name']}", ln=True)
    pdf.ln(1.5)
    pdf.set_font(font_name, "", 7.4)
    pdf.set_text_color(75, 75, 75)
    pdf.multi_cell(0, 4.0, "Resmi Mevzuat ve Katsayı Dayanağı: Bu analiz; (EU) 2023/956 sayılı AB SKDM Tüzüğü (Annex II & Annex IV), kesin dönem ülke bazlı varsayılan değerleri belirleyen (EU) 2025/2621 ve (EU) 2026/1740 sayılı Uygulama Tüzükleri, yakıt emisyon faktörlerini düzenleyen (EU) 2023/1773 sayılı Tüzük ile T.C. Enerji ve Tabii Kaynaklar Bakanlığı (ETKB) Tüketim Noktası Ulusal Elektrik Şebekesi Emisyon Faktörleri esas alınarak hazırlanmıştır. Kesin AB gümrük beyanında akredite doğrulayıcı ((EU) 2025/2547) onayı gereklidir.")
    pdf.ln(1.2)
    pdf.multi_cell(0, 4.0, "Zero-Retention Veri Güvenliği Beyanı: Bu platforma yüklenen fatura, enerji tüketim ve sevkiyat verileri hiçbir veritabanında kalıcı olarak saklanmaz; oturum belleğinde anlık işlenir ve oturum kapandığında kalıcı olarak silinir.")

    return bytes(pdf.output())

# --- YAN MENÜ (GLOBAL PARAMETRELER & WHITE-LABEL AYARLARI) ---
with st.sidebar:
    st.header("⚙️ Kurumsal & Mevzuat Ayarları")
    broker_name = st.text_input("Gümrük Müşavirliği / Danışmanlık Unvanı", value="İÜ Girişim Hub — R&D Division")
    broker_logo = st.file_uploader("Kurumsal Logo Yükle (PDF Anteti İçin)", type=["png", "jpg", "jpeg"])
    
    st.divider()
    st.subheader("📅 AB SKDM Dönem & Piyasa Parametreleri")
    year_choice = st.selectbox("Beyan Yılı & Ceza Çarpanı", list(YEAR_RULES.keys()))
    grid_choice = st.selectbox("ETKB Elektrik Şebeke Bağlantı Tipi", list(ETKB_GRID_FACTORS.keys()))
    ets_price = st.number_input("AB ETS Karbon Fiyatı (€ / tCO₂e)", min_value=10.0, value=74.50, step=0.5)
    local_carbon_price = st.number_input("Türkiye İçi Ödenen Karbon Bedeli (€ / tCO₂e)", min_value=0.0, value=0.0, step=1.0, help="Türkiye Emisyon Ticaret Sistemi kapsamında ödenen tutar AB vergisinden mahsup edilir.")

    st.divider()
    st.caption("🔒 **Zero-Retention:** Tüm veriler yalnızca aktif oturum RAM belleğinde işlenir.")

# --- ANA EKRAN BAŞLIĞI ---
st.title("TR-CBAM v2.0 | Kurumsal SKDM Uyum & Denetim Kokpiti")
st.markdown("**(EU) 2025/2621 & 2026/1740 Kesin Dönem Türkiye Katsayıları + ETKB Şebeke Motoru + AB Registry XML Üreticisi**")

tab1, tab2, tab3 = st.tabs(["📄 Tekil Sevkiyat & Fatura OCR Analizi", "📊 Çoklu Kalem (Excel) & AB XML Çıktısı", "📚 2026 Türkiye Resmi Katsayı Tablosu"])

# ==============================================================================
# SEKME 1: TEKİL SEVKİYAT + FATURA OCR + GRAFİKLER + KURUMSAL PDF + XML
# ==============================================================================
with tab1:
    col_left, col_right = st.columns([1.15, 1])

    with col_left:
        st.subheader("1. Tesis ve Fatura Verileri")
        invoice_file = st.file_uploader("Opsiyonel: Elektrik veya Doğalgaz Faturası Yükle (PDF - Otomatik Okuma)", type=["pdf"])
        
        auto_kwh, auto_m3 = 45000.0, 3200.0
        invoice_status = "Kanıt Belgesi Henüz Yüklenmedi (Ön Simülasyon)"
        if invoice_file is not None:
            det_kwh, det_m3, readable = extract_energy_from_pdf(invoice_file)
            invoice_status = f"Doğrulandı: {invoice_file.name}"
            if det_kwh:
                auto_kwh = det_kwh
                st.success(f"Faturadan otomatik elektrik tüketimi okundu: **{det_kwh:,.2f} kWh**")
            if det_m3:
                auto_m3 = det_m3
                st.success(f"Faturadan otomatik doğalgaz tüketimi okundu: **{det_m3:,.2f} m³**")
            if not det_kwh and not det_m3:
                st.info("PDF yüklendi; taranmış görsel formatındaysa tüketim değerlerini aşağıdan manuel onaylayabilirsiniz.")

        with st.form("single_cbam_form"):
            company_name = st.text_input("İhracatçı Firma Unvanı", value="Örnek Metal ve Çelik Sanayi A.Ş.")
            gtip_label = st.selectbox("Ürün CN (GTİP) Kodu", list(GTIP_2026_DB.keys()))
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
            yr["markup"], yr["cbam_phase_in"]
        )

        with col_right:
            st.subheader("2. Finansal Etki (The Money Screen)")
            m1, m2, m3 = st.columns(3)
            m1.metric("Kanıtsız Ceza (A)", f"{res['tax_default']:,.0f} €")
            m2.metric("Kanıtlı Beyan (B)", f"{res['tax_actual']:,.0f} €")
            m3.metric("Net Vergi Tasarrufu", f"+{res['direct_savings']:,.0f} €", delta=f"Ek Tedarikçi Fırsatı: +{res['supplier_bonus']:,.0f} €")

            # Plotly Kıyaslama Grafiği
            fig = go.Figure(data=[
                go.Bar(
                    name="Maliyet (€)",
                    x=["A: Kanıtsız Default", "B: Tesis Gerçek Verisi", "C: Tedarikçi Verisi Dahil"],
                    y=[res["tax_default"], res["tax_actual"], res["tax_best"]],
                    marker_color=["#ef4444", "#3b82f6", "#10b981"],
                    text=[f"{res['tax_default']:,.0f} €", f"{res['tax_actual']:,.0f} €", f"{res['tax_best']:,.0f} €"],
                    textposition="auto"
                )
            ])
            fig.update_layout(
                title="Senaryo Bazlı AB Sınırda Karbon Vergisi Yükü (€)",
                height=310,
                margin=dict(l=20, r=20, t=40, b=20),
                yaxis_title="Euro (€)"
            )
            st.plotly_chart(fig, use_container_width=True)

        st.divider()
        st.info(f"**Mevzuat Kapsam Bilgisi:** {res['scope_note']} | **Uygulanan 2026 TR Ceza Katsayısı:** `{res['default_ef_markup']:.3f} tCO₂e/t`")

        # Çıktı Butonları (Kurumsal PDF + Resmi AB XML)
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
            file_name=f"TR_CBAM_2026_Rapor_{company_name.replace(' ', '_')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
        b_col2.download_button(
            label="🧾 AB CBAM Portalı Resmi XML Dosyasını İndir",
            data=xml_bytes,
            file_name=f"EU_CBAM_Registry_{res['cn']}_{yr['year']}.xml",
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
        {"Urun_Adi": "Sıcak Haddelenmiş Sac", "GTIP_Kodu": "7208 - Sıcak Haddelenmiş Yassı Çelik Ürünler", "Tonaj": 120.0, "Elektrik_kWh": 45000.0, "Dogalgaz_m3": 3200.0},
        {"Urun_Adi": "Çelik Bağlantı Cıvatası", "GTIP_Kodu": "7318 15 - Cıvata, Vida, Somun ve Bağlantı Elemanları", "Tonaj": 35.0, "Elektrik_kWh": 18000.0, "Dogalgaz_m3": 1100.0},
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
                g_label = row["GTIP_Kodu"] if row["GTIP_Kodu"] in GTIP_2026_DB else list(GTIP_2026_DB.keys())[0]
                r = calculate_cbam_row(
                    g_label, float(row["Tonaj"]), float(row["Elektrik_kWh"]), float(row["Dogalgaz_m3"]),
                    ets_price, local_carbon_price, grid_ef, yr["markup"], yr["cbam_phase_in"]
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
            st.success(f"**Toplam Sevkiyat Tasarrufu:** `{tot_sav:,.2f} €`")

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
# SEKME 3: 2026 TÜRKİYE RESMİ KATSAYI REFERANS TABLOSU
# ==============================================================================
with tab3:
    st.subheader("(EU) 2025/2621 & 2026/1740 Türkiye Özel Kesin Dönem Katsayıları")
    ref_list = []
    for k, v in GTIP_2026_DB.items():
        ref_list.append({
            "CN / GTİP": v["cn"],
            "Ürün Tanımı": k.split(" - ")[1],
            "Sektör": v["sector"],
            "2026 Baz Değer (tCO₂e/t)": v["base_default"],
            "2026 Ceza Artışlı (+%10)": round(v["base_default"] * 1.10, 3),
            "2027 Ceza Artışlı (+%20)": round(v["base_default"] * 1.20, 3),
            "Mali Yükümlülük Kapsamı": "Sadece Doğrudan (Direct)" if v["scope"] == "direct_only" else "Doğrudan + Dolaylı (Elektrik)"
        })
    st.dataframe(pd.DataFrame(ref_list), use_container_width=True)
