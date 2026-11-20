import streamlit as st
from fpdf import FPDF
from datetime import datetime
import os

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="TR-CBAM | Pre-Audit Uyum Kokpiti", layout="centered")

# --- DETERMINİSTİK EMİSYON VE BENCHMARK VERİ TABANI (AB CBAM & TR ŞEBEKE) ---
TR_GRID_EF = 0.442  # tCO2 / MWh (Türkiye Ulusal Elektrik Şebekesi Emisyon Faktörü)
NAT_GAS_EF = 0.00192  # tCO2 / m3 (Doğalgaz Yanma Emisyon Faktörü)

GTIP_BENCHMARKS = {
    "7208 - Sıcak Haddelenmiş Yassı Çelik Ürünler": {"default_ef": 2.15, "precursor_ef": 1.35},
    "7318 - Demir/Çelik Cıvata, Vida, Somun ve Bağlantı Elemanları": {"default_ef": 2.65, "precursor_ef": 1.85},
    "7604 - Alüminyum Çubuk ve Profiller": {"default_ef": 8.40, "precursor_ef": 6.20},
    "7606 - Alüminyum Sac, Levha ve Şeritler": {"default_ef": 8.95, "precursor_ef": 6.50},
    "2523 - Portland Çimento ve Klinker": {"default_ef": 0.92, "precursor_ef": 0.00},
}

# --- TÜRKÇE UNICODE FONT YÜKLEYİCİ (WINDOWS & LINUX UYUMLU) ---
def setup_unicode_font(pdf):
    win_arial = "C:/Windows/Fonts/arial.ttf"
    win_arial_bold = "C:/Windows/Fonts/arialbd.ttf"
    win_arial_italic = "C:/Windows/Fonts/ariali.ttf"
    
    if os.path.exists(win_arial):
        pdf.add_font("TrFont", "", win_arial)
        pdf.add_font("TrFont", "B", win_arial_bold if os.path.exists(win_arial_bold) else win_arial)
        pdf.add_font("TrFont", "I", win_arial_italic if os.path.exists(win_arial_italic) else win_arial)
        return "TrFont"
    return "Helvetica"

# --- PDF RAPOR ÜRETİCİ ("THE MONEY SCREEN" İLK SAYFADA) ---
def generate_pre_audit_pdf(data):
    pdf = FPDF()
    pdf.add_page()
    font_name = setup_unicode_font(pdf)
    
    # Üst Kurumsal Başlık Şeridi
    pdf.set_fill_color(15, 23, 42)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(font_name, "B", 14)
    pdf.cell(0, 12, "TR-CBAM PRE-AUDIT & SINIR VERGİSİ ANALİZ RAPORU", ln=True, align="C", fill=True)
    
    pdf.set_text_color(80, 80, 80)
    pdf.set_font(font_name, "", 9)
    pdf.cell(0, 7, f"Rapor Tarihi: {datetime.now().strftime('%d.%m.%Y %H:%M')}  |  Referans No: CBAM-{int(datetime.now().timestamp())}", ln=True, align="C")
    pdf.ln(5)

    # 1. Sevkiyat ve Tesis Bilgileri
    pdf.set_text_color(0, 0, 0)
    pdf.set_font(font_name, "B", 11)
    pdf.cell(0, 8, "1. SEVKİYAT VE TESİS BİLGİLERİ", ln=True)
    pdf.set_font(font_name, "", 10)
    pdf.cell(0, 6, f"İhracatçı Firma Unvanı: {data['company_name']}", ln=True)
    pdf.cell(0, 6, f"GTİP / Ürün Grubu: {data['gtip']}", ln=True)
    pdf.cell(0, 6, f"AB Sevkiyat Miktarı: {data['tonnage']:,.2f} Ton   |   AB ETS Karbon Fiyatı: {data['ets_price']:.2f} EUR / tCO2", ln=True)
    pdf.cell(0, 6, f"Denetim Kanıt Belgesi: {data['invoice_status']}", ln=True)
    pdf.ln(6)

    # 2. THE MONEY SCREEN (KOMİTE KURAL #2)
    pdf.set_fill_color(241, 245, 249)
    pdf.set_font(font_name, "B", 11)
    pdf.cell(0, 9, "2. FİNANSAL ETKİ KIYASLAMASI (THE MONEY SCREEN)", ln=True, fill=True)
    pdf.ln(2)

    # Tablo Başlıkları
    pdf.set_font(font_name, "B", 9.5)
    pdf.cell(95, 8, "Beyan Senaryosu", border=1)
    pdf.cell(45, 8, "Gömülü Emisyon", border=1, align="C")
    pdf.cell(50, 8, "Sınır Vergisi (CBAM)", border=1, align="R", ln=True)

    # Senaryo A
    pdf.set_font(font_name, "", 9.5)
    pdf.cell(95, 8, "Senaryo A: Kanıtsız Beyan (AB Varsayılan Değer)", border=1)
    pdf.cell(45, 8, f"{data['default_total_em']:,.2f} tCO2", border=1, align="C")
    pdf.cell(50, 8, f"{data['tax_scenario_a']:,.2f} EUR", border=1, align="R", ln=True)

    # Senaryo B
    pdf.cell(95, 8, "Senaryo B: Kanıtlı Gerçek Tesis Enerji Verisi", border=1)
    pdf.cell(45, 8, f"{data['actual_total_em']:,.2f} tCO2", border=1, align="C")
    pdf.cell(50, 8, f"{data['tax_scenario_b']:,.2f} EUR", border=1, align="R", ln=True)

    # Net Tasarruf Satırı
    pdf.set_font(font_name, "B", 10.5)
    pdf.set_fill_color(220, 252, 231)
    pdf.cell(140, 10, "NET SINIR VERGİSİ TASARRUFU (SEVKİYAT BAŞINA)", border=1, fill=True)
    pdf.cell(50, 10, f"+{data['net_savings']:,.2f} EUR", border=1, align="R", ln=True, fill=True)
    pdf.ln(6)

    # 3. Emisyon Dağılımı ve Tedarikçi Baskı Uyarısı
    pdf.set_font(font_name, "B", 11)
    pdf.cell(0, 8, "3. EMİSYON DAĞILIMI VE TEDARİKÇİ UYARISI", ln=True)
    pdf.set_font(font_name, "", 10)
    pdf.cell(0, 6, f"• Scope 1 (Doğalgaz Kaynaklı Doğrudan Emisyon): {data['scope1']:,.2f} tCO2", ln=True)
    pdf.cell(0, 6, f"• Scope 2 (Elektrik Kaynaklı Dolaylı Emisyon): {data['scope2']:,.2f} tCO2", ln=True)
    pdf.cell(0, 6, f"• Precursor (Öncül Hammadde Varsayılan Emisyonu): {data['precursor_total']:,.2f} tCO2", ln=True)
    pdf.ln(3)
    
    pdf.set_fill_color(254, 243, 199)
    pdf.set_font(font_name, "I", 9)
    pdf.multi_cell(0, 5.5, f"TEDARİKÇİ BASKI NOTU: Hammadde tedarikçiniz gerçek emisyon verisi paylaşmadığı için öncül madde (precursor) kaleminde AB varsayılan katsayısı uygulanmıştır. Tedarikçiden doğrulanmış veri alınması halinde sevkiyat başına ek {data['precursor_penalty']:,.2f} EUR daha sınır vergisi avantajı sağlanabilir.", border=1, fill=True)
    pdf.ln(8)

    # 4. Gümrük Müşaviri / Danışman Kaşe ve Onay Alanı
    pdf.set_font(font_name, "B", 9.5)
    pdf.cell(0, 6, f"Raporu Düzenleyen Kurum / Gümrük Müşavirliği: {data['broker_name']}", ln=True)
    pdf.set_font(font_name, "", 8.5)
    pdf.set_text_color(100, 100, 100)
    pdf.multi_cell(0, 4.5, "Yasal Uyarı: Bu belge AB (EU) 2023/956 sayılı CBAM Tüzüğü kapsamında ön değerlendirme ve akredite doğrulayıcı (TÜRKAK/AB Verifier) hazırlık dosyası olarak oluşturulmuştur.")

    return bytes(pdf.output())

# --- ARAYÜZ (TEK SAYFA, TEK AKIŞ - NO AUTH) ---
st.title("TR-CBAM | Pre-Audit & Sınır Vergisi Kokpiti")
st.caption("Gümrük Müşavirleri ve İhracatçılar İçin 60 Saniyede Kanıtlı CBAM Maliyet Simülasyonu")

with st.form("cbam_form"):
    col1, col2 = st.columns(2)
    with col1:
        company_name = st.text_input("İhracatçı Firma Unvanı", value="Örnek Metal Sanayi A.Ş.")
        gtip = st.selectbox("Ürün GTİP Kodu", list(GTIP_BENCHMARKS.keys()))
        tonnage = st.number_input("AB Sevkiyat Miktarı (Ton)", min_value=1.0, value=120.0, step=10.0)
    with col2:
        elec_kwh = st.number_input("Aylık Elektrik Tüketimi (kWh)", min_value=0.0, value=45000.0, step=1000.0)
        gas_m3 = st.number_input("Aylık Doğalgaz Tüketimi (m³)", min_value=0.0, value=3200.0, step=100.0)
        ets_price = st.number_input("Güncel AB ETS Karbon Fiyatı (€ / tCO2)", min_value=10.0, value=72.5, step=1.0)

    broker_name = st.text_input("Raporu Düzenleyen Gümrük Müşavirliği / Danışmanlık", value="İÜ Girişim Hub — R&D Division")
    invoice_file = st.file_uploader("Kanıt Belgesi: Elektrik / Doğalgaz Faturası (PDF/JPG)", type=["pdf", "jpg", "png"])
    submitted = st.form_submit_button("60 Saniyede Pre-Audit Analizini Üret", use_container_width=True)

if submitted:
    bench = GTIP_BENCHMARKS[gtip]
    scope1 = gas_m3 * NAT_GAS_EF
    scope2 = (elec_kwh / 1000.0) * TR_GRID_EF
    precursor_total = tonnage * bench["precursor_ef"]
    
    actual_total_em = scope1 + scope2 + precursor_total
    default_total_em = tonnage * bench["default_ef"]
    
    tax_scenario_a = default_total_em * ets_price
    tax_scenario_b = actual_total_em * ets_price
    net_savings = max(0.0, tax_scenario_a - tax_scenario_b)
    precursor_penalty = precursor_total * ets_price * 0.35

    st.divider()
    st.subheader("Finansal Etki Kıyaslaması (The Money Screen)")
    
    m1, m2, m3 = st.columns(3)
    m1.metric("Senaryo A (Kanıtsız AB Default)", f"{tax_scenario_a:,.0f} €")
    m2.metric("Senaryo B (Kanıtlı Tesis Verisi)", f"{tax_scenario_b:,.0f} €")
    m3.metric("Net Sınır Vergisi Tasarrufu", f"+{net_savings:,.0f} €", delta="Sevkiyat Başı Kazanç")

    if precursor_total > 0:
        st.warning(f"**Tedarikçi Baskı Uyarısı:** Hammadde tedarikçiniz gerçek emisyon verisini paylaşmadığı için AB varsayılan katsayısı uygulandı. Tedarikçiden veri alınırsa sevkiyat başına ek **~{precursor_penalty:,.0f} €** daha vergi avantajı doğabilir.")

    pdf_data = {
        "company_name": company_name,
        "gtip": gtip,
        "tonnage": tonnage,
        "ets_price": ets_price,
        "broker_name": broker_name,
        "invoice_status": f"Yüklendi ({invoice_file.name})" if invoice_file else "Kanıt Bekleniyor (Taslak Beyan)",
        "default_total_em": default_total_em,
        "actual_total_em": actual_total_em,
        "tax_scenario_a": tax_scenario_a,
        "tax_scenario_b": tax_scenario_b,
        "net_savings": net_savings,
        "scope1": scope1,
        "scope2": scope2,
        "precursor_total": precursor_total,
        "precursor_penalty": precursor_penalty,
    }
    
    pdf_bytes = generate_pre_audit_pdf(pdf_data)
    st.download_button(
        label="Pre-Audit PDF Raporunu İndir (Müşteri Sunumuna Hazır)",
        data=pdf_bytes,
        file_name=f"TR_CBAM_PreAudit_{company_name.replace(' ', '_')}.pdf",
        mime="application/pdf",
        use_container_width=True
    )