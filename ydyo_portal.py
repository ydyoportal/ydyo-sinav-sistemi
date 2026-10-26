import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import psycopg2
import json
import io
import hashlib
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Flowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="KTO Karatay YDY Portalı", page_icon="🔐", layout="wide")

# ==========================================
# GÜVENLİK VE VERİTABANI BAĞLANTISI (SUPABASE)
# ==========================================
SALT = "KTO_YDYO_SECURE_SALT_2026"

def hash_sifre(sifre: str) -> str:
    if not sifre: return ""
    if len(sifre) == 64 and all(c in '0123456789abcdefABCDEF' for c in sifre):
        return sifre
    return hashlib.sha256((sifre + SALT).encode('utf-8')).hexdigest()

def sifre_dogrula(girilen_sifre: str, hashli_sifre: str) -> bool:
    return hash_sifre(girilen_sifre) == hashli_sifre

def get_db_connection():
    # Streamlit Cloud'da 'secrets' üzerinden okur, yoksa doğrudan fallback URI kullanır
    if "SUPABASE_URI" in st.secrets:
        uri = st.secrets["SUPABASE_URI"]
    else:
        uri = "postgresql://postgres.atyewikzyuvembeikirx:X9R%24t9vVu7ctF%24%21@aws-0-ap-southeast-2.pooler.supabase.com:5432/postgres"
    return psycopg2.connect(uri)

# ==========================================
# PDF MOTORU VE VEKTÖREL ÇİZİMLER
# ==========================================
def tr_to_eng(text):
    if not isinstance(text, str): return str(text)
    tr_map = {'ı':'i', 'i':'i', 'I':'I', 'İ':'I', 'ğ':'g', 'Ğ':'G', 'ü':'u', 'Ü':'U', 'ş':'s', 'Ş':'S', 'ö':'o', 'Ö':'O', 'ç':'c', 'Ç':'C'}
    for k, v in tr_map.items(): text = text.replace(k, v)
    return text

class ModernBar(Flowable):
    def __init__(self, width, height, percentage):
        Flowable.__init__(self)
        self.width = width
        self.height = height
        self.percentage = max(0, min(100, percentage))
        if self.percentage >= 75: self.color = colors.HexColor("#27AE60")
        elif self.percentage >= 50: self.color = colors.HexColor("#F1C40F")
        else: self.color = colors.HexColor("#E74C3C")

    def draw(self):
        self.canv.setFillColor(colors.HexColor("#EAECEE"))
        self.canv.roundRect(0, 0, self.width, self.height, 3, fill=1, stroke=0)
        fill_width = (self.width * self.percentage) / 100
        if fill_width > 0:
            self.canv.setFillColor(self.color)
            self.canv.roundRect(0, 0, fill_width, self.height, 3, fill=1, stroke=0)

styles = getSampleStyleSheet()
header_style = ParagraphStyle('Header', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=14, textColor=colors.HexColor("#1A5276"), alignment=1, spaceAfter=5)
sub_header_style = ParagraphStyle('SubHeader', parent=styles['Normal'], fontName='Helvetica', fontSize=10, textColor=colors.grey, alignment=1, spaceAfter=20)
title_style = ParagraphStyle('ModernTitle', parent=styles['Title'], fontName='Helvetica-Bold', fontSize=16, textColor=colors.HexColor("#2C3E50"), spaceAfter=10)
info_style = ParagraphStyle('Info', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor("#34495E"), spaceAfter=15)
cell_style = ParagraphStyle('Cell', parent=styles['Normal'], fontName='Helvetica', fontSize=9, alignment=1)
cell_style_left = ParagraphStyle('CellLeft', parent=styles['Normal'], fontName='Helvetica', fontSize=9, alignment=0)
header_cell_style = ParagraphStyle('HeaderCell', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=colors.whitesmoke, alignment=1)

def build_pdf_doc(elements, is_landscape=False):
    buffer = io.BytesIO()
    page_format = landscape(A4) if is_landscape else A4
    doc = SimpleDocTemplate(buffer, pagesize=page_format, rightMargin=30, leftMargin=30, topMargin=40, bottomMargin=30)
    doc.build(elements)
    buffer.seek(0)
    return buffer

def create_landscape_class_pdf(veri):
    elements = []
    elements.append(Paragraph("KTO KARATAY UNIVERSITY", header_style))
    elements.append(Paragraph("School of Foreign Languages | Testing Office", sub_header_style))
    elements.append(Paragraph(tr_to_eng(veri['baslik']), title_style))
    elements.append(Paragraph(tr_to_eng(veri['alt_baslik']), info_style))
    
    if veri.get('yorum'):
        elements.append(Paragraph("<b>Instructor / Admin Comments:</b>", styles['Heading4']))
        elements.append(Paragraph(tr_to_eng(veri['yorum']).replace('\n', '<br />'), styles['Normal']))
        elements.append(Spacer(1, 15))
        
    kpi_df = veri['df_kpi']
    kpi_data = [[Paragraph("Metric", header_cell_style), Paragraph("Value", header_cell_style)]]
    for _, row in kpi_df.iterrows(): kpi_data.append([Paragraph(tr_to_eng(str(row['Metric'])), cell_style_left), Paragraph(str(row['Value']), cell_style)])
        
    kpi_table = Table(kpi_data, colWidths=[130, 100])
    kpi_table.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2C3E50")), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BOTTOMPADDING', (0,0), (-1,-1), 8), ('TOPPADDING', (0,0), (-1,-1), 8), ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey), ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#F8F9F9"))]))
    
    skills_data = [[Paragraph("Section / Part", header_cell_style), Paragraph("Avg", header_cell_style), Paragraph("Max", header_cell_style), Paragraph("Performance", header_cell_style)]]
    for b_dict in veri['beceri_listesi']:
        bar = ModernBar(140, 10, b_dict['yuzde'])
        skills_data.append([Paragraph(tr_to_eng(b_dict['isim']), cell_style_left), Paragraph(f"{b_dict['ort']:.1f}", cell_style), Paragraph(str(b_dict['max']), cell_style), bar])
        
    skill_table = Table(skills_data, colWidths=[200, 50, 50, 160])
    s_style = [('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2C3E50")), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey)]
    for i in range(1, len(skills_data)):
        if i % 2 == 0: s_style.append(('BACKGROUND', (0,i), (-1,i), colors.HexColor("#F8F9F9")))
    skill_table.setStyle(TableStyle(s_style))
    
    kapsayici_tablo = Table([[kpi_table, skill_table]], colWidths=[250, 532])
    kapsayici_tablo.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    elements.extend([kapsayici_tablo, Spacer(1, 20)])
    
    df_liste = veri['df_liste'].copy()
    elements.append(Paragraph("<b>Detailed Student List & Scores</b>", styles['Heading4']))
    elements.append(Spacer(1, 5))
    
    col_count = len(df_liste.columns)
    base_widths = {'Student ID': 70, 'Name Surname': 160, 'Total Score': 60}
    kalan_genislik = 782 - sum(base_widths.values())
    skill_col_width = kalan_genislik / max(1, col_count - 3)
    col_widths = [base_widths.get(c, skill_col_width) for c in df_liste.columns]
    
    tablo_verisi = [[Paragraph(c, header_cell_style) for c in df_liste.columns]]
    for _, row in df_liste.iterrows():
        row_data = []
        for col_idx, c_name in enumerate(df_liste.columns):
            c_text = tr_to_eng(str(row[c_name]))
            c_style = cell_style_left if c_name == 'Name Surname' else cell_style
            row_data.append(Paragraph(c_text, c_style))
        tablo_verisi.append(row_data)
        
    l_table = Table(tablo_verisi, colWidths=col_widths, repeatRows=1)
    l_style = [('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2C3E50")), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey), ('BOTTOMPADDING', (0,0), (-1,0), 8), ('TOPPADDING', (0,0), (-1,0), 8)]
    for i in range(1, len(tablo_verisi)):
        if i % 2 == 0: l_style.append(('BACKGROUND', (0,i), (-1,i), colors.HexColor("#F8F9F9")))
    l_table.setStyle(TableStyle(l_style))
    elements.append(l_table)
    
    return build_pdf_doc(elements, is_landscape=True)

def create_modern_pdf(rapor_verileri):
    elements = []
    for idx, veri in enumerate(rapor_verileri):
        elements.append(Paragraph("KTO KARATAY UNIVERSITY", header_style))
        elements.append(Paragraph("School of Foreign Languages | Testing Office", sub_header_style))
        elements.append(Paragraph(tr_to_eng(veri['baslik']), title_style))
        if veri.get('alt_baslik'): elements.append(Paragraph(tr_to_eng(veri['alt_baslik']), info_style))
        
        if veri.get('yorum'):
            elements.append(Paragraph("<b>Instructor / Admin Comments:</b>", styles['Heading4']))
            elements.append(Paragraph(tr_to_eng(veri['yorum']).replace('\n', '<br />'), styles['Normal']))
            elements.append(Spacer(1, 15))
            
        if veri.get('beceri_grafik_verisi'):
            elements.append(Paragraph("<b>Skill Profile</b>", styles['Heading4']))
            elements.append(Spacer(1, 5))
            b_data = []
            for b_name, b_val in veri['beceri_grafik_verisi'].items():
                bar = ModernBar(120, 10, b_val)
                b_data.append([Paragraph(tr_to_eng(b_name), cell_style_left), bar, Paragraph(f"%{b_val:.1f}", cell_style)])
            if b_data:
                b_table = Table(b_data, colWidths=[120, 130, 50])
                b_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BOTTOMPADDING', (0,0), (-1,-1), 4), ('TOPPADDING', (0,0), (-1,-1), 4)]))
                elements.extend([b_table, Spacer(1, 15)])
            
        df_listesi = veri.get('df_list', [])
        for tablo_baslik, df_tablo in df_listesi:
            if not df_tablo.empty:
                if tablo_baslik:
                    elements.append(Paragraph(f"<b>{tr_to_eng(tablo_baslik)}</b>", styles['Heading4']))
                    elements.append(Spacer(1, 5))
                    
                df_temp = df_tablo.copy()
                df_temp.columns = [tr_to_eng(c) for c in df_temp.columns]
                
                cols_count = len(df_temp.columns)
                if cols_count == 6: col_widths = [75, 60, 210, 50, 60, 80]
                elif cols_count == 3: col_widths = [100, 335, 100]
                elif cols_count == 2: col_widths = [267, 268]
                else: col_widths = [535 / cols_count] * cols_count
                
                tablo_verisi = [[Paragraph(c, header_cell_style) for c in df_temp.columns]]
                for _, row in df_temp.iterrows():
                    row_data = []
                    for col_idx, cell_val in enumerate(row):
                        c_text = tr_to_eng(str(cell_val))
                        c_style = cell_style_left if cols_count == 6 and col_idx == 2 else cell_style
                        row_data.append(Paragraph(c_text, c_style))
                    tablo_verisi.append(row_data)
                    
                tablo = Table(tablo_verisi, colWidths=col_widths, repeatRows=1)
                t_style = [('BACKGROUND', (0,0), (-1,0), colors.HexColor("#283747")), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BOTTOMPADDING', (0,0), (-1,0), 8), ('TOPPADDING', (0,0), (-1,0), 8), ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey)]
                for i in range(1, len(tablo_verisi)):
                    if i % 2 == 0: t_style.append(('BACKGROUND', (0,i), (-1,i), colors.HexColor("#F8F9F9")))
                
                if "Basari (%)" in df_temp.columns:
                    perc_col_idx = df_temp.columns.get_loc("Basari (%)")
                    for i in range(1, len(tablo_verisi)):
                        try:
                            val = float(str(df_temp.iloc[i-1]["Basari (%)"]).replace('%', '').strip())
                            if val >= 75: t_style.append(('BACKGROUND', (perc_col_idx, i), (perc_col_idx, i), colors.HexColor("#D5F5E3")))
                            elif val >= 50: t_style.append(('BACKGROUND', (perc_col_idx, i), (perc_col_idx, i), colors.HexColor("#FCF3CF")))
                            else: t_style.append(('BACKGROUND', (perc_col_idx, i), (perc_col_idx, i), colors.HexColor("#FADBD8")))
                        except: pass
                tablo.setStyle(TableStyle(t_style))
                elements.extend([tablo, Spacer(1, 20)])
            
        if idx < len(rapor_verileri) - 1: elements.append(PageBreak())
            
    return build_pdf_doc(elements, is_landscape=False)

# ==========================================
# GİRİŞ KAPISI VE OTURUM
# ==========================================
if 'giris_basarili' not in st.session_state:
    st.session_state.update({'giris_basarili': False, 'kullanici_adi': "", 'rol': "", 'isim': ""})

def giris_yap(k_adi, sifre):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT sifre, rol, isim FROM kullanicilar WHERE kullanici_adi=%s", (k_adi,))
    kullanici = cur.fetchone()
    cur.close()
    conn.close()
    if kullanici and sifre_dogrula(sifre, kullanici[0]):
        st.session_state.update({'giris_basarili': True, 'kullanici_adi': k_adi, 'rol': kullanici[1], 'isim': kullanici[2]})
        st.rerun()
    else:
        st.error("Kullanıcı adı veya şifre hatalı!")

def cikis_yap():
    st.session_state.update({'giris_basarili': False, 'kullanici_adi': "", 'rol': "", 'isim': ""})
    st.rerun()

if not st.session_state['giris_basarili']:
    st.title("🔐 KTO Karatay YDY - Merkezi Giriş")
    with st.form("giris_formu"):
        k_adi = st.text_input("Kullanıcı Adı")
        sifre = st.text_input("Şifre", type="password")
        if st.form_submit_button("Sisteme Giriş Yap", type="primary"):
            giris_yap(k_adi, sifre)
else:
    st.sidebar.title(f"👤 {st.session_state['isim']}")
    st.sidebar.info(f"Yetki: {st.session_state['rol']}")
    if st.sidebar.button("🚪 Güvenli Çıkış"):
        cikis_yap()
        
    st.title("🎓 KTO Karatay YDY Öğrenme Analitiği Sistemi")

    if st.session_state['rol'] in ["Admin", "Testing"]:
        if st.session_state['rol'] == "Admin":
            tabs = st.tabs(["📝 Sınav Tanımla", "🗑️ Sınav Yönetimi", "📋 Öğrenci & Atamalar", "👥 Kullanıcılar", "📊 Analitik Raporlar"])
            tab_sinav, tab_yonetim, tab_atamalar, tab_kullanici, tab_rapor = tabs[0], tabs[1], tabs[2], tabs[3], tabs[4]
        else:
            tabs = st.tabs(["📝 Sınav Tanımla", "🗑️ Sınav Yönetimi", "📊 Analitik Raporlar"])
            tab_sinav, tab_yonetim, tab_rapor = tabs[0], tabs[1], tabs[2]
        
        with tab_sinav:
            st.subheader("Standart Sınav Kodu Oluşturucu")
            c1, c2, c3, c4 = st.columns(4)
            with c1: secilen_modul = st.selectbox("Modül", ["M1", "M2", "M3", "M4", "YazOkulu"])
            with c2: secilen_kur = st.selectbox("Kur Seviyesi", ["A1", "A2", "B1", "B2", "PF"])
            with c3: secilen_tur = st.selectbox("Sınav Türü", ["Quiz_1", "Quiz_2", "Midterm", "Final"])
            with c4: tolerans_baremi = st.number_input("Çift Okuma Tolerans Farkı", min_value=0, value=6)
            sinav_kodu = f"{secilen_modul}_{secilen_kur}_{secilen_tur}"
            
            st.subheader("Gelişmiş Belirtke Tablosu (ToS) İnşaatçısı")
            tos_file = st.file_uploader("📥 Excel'den Hazır ToS Yükle (Sütunlar: Ana Bölüm, Alt Bölüm (Part), Ölçülen Kazanımlar, Max Puan)", type=["xlsx", "xls"])
            if tos_file:
                try:
                    df_up_tos = pd.read_excel(tos_file)
                    st.session_state['taslak_df'] = df_up_tos
                    st.success("ToS başarıyla Excel'den içeri aktarıldı!")
                except: st.error("Excel dosya formatı uyumsuz.")

            ANA_BOLUMLER = [
                "Listening I", "Listening II", "Reading I", "Reading II", 
                "Grammar", "Functional Language", "Use of English", "Vocabulary", 
                "Speaking", "Speaking 1st", "Speaking 2nd", "Speaking 3rd", 
                "Writing", "Writing (1. Okuma)", "Writing (2. Okuma)", "Writing (3. Okuma)"
            ]
            if 'taslak_df' not in st.session_state:
                st.session_state['taslak_df'] = pd.DataFrame({"Ana Bölüm": ["Grammar", "Writing (1. Okuma)", "Writing (2. Okuma)"], "Alt Bölüm (Part)": ["Part A", "Task 1", "Task 1"], "Ölçülen Kazanımlar": ["Present Simple", "Coherence", "Coherence"], "Max Puan": [15, 25, 25]})
            
            st.caption("💡 İpucu: Excel'deki verilerinizi kopyalayarak doğrudan aşağıdaki tablonun içine yapıştırabilirsiniz.")
            duzenlenen_rubrik = st.data_editor(st.session_state['taslak_df'], num_rows="dynamic", column_config={"Ana Bölüm": st.column_config.SelectboxColumn("Ana Bölüm", options=ANA_BOLUMLER, required=True)}, use_container_width=True)
            
            if st.button("🚀 Sınavı Onayla & Öğretim Elemanlarına Aç", type="primary"):
                conn = get_db_connection()
                cur = conn.cursor()
                cur.execute("SELECT sinav_kodu FROM sinavlar WHERE sinav_kodu=%s", (sinav_kodu,))
                if cur.fetchone(): st.warning("Bu sınav kodu zaten açık!")
                else:
                    cur.execute("INSERT INTO sinavlar (sinav_kodu) VALUES (%s)", (sinav_kodu,))
                    cur.execute("INSERT INTO sinav_ayarlari (sinav_kodu, tolerans) VALUES (%s, %s)", (sinav_kodu, tolerans_baremi))
                    for _, row in duzenlenen_rubrik.iterrows():
                        cur.execute("INSERT INTO sinav_tos (sinav_kodu, ana_bolum, alt_bolum, kazanimlar, max_puan) VALUES (%s, %s, %s, %s, %s)", 
                                    (sinav_kodu, row['Ana Bölüm'], row['Alt Bölüm (Part)'], row['Ölçülen Kazanımlar'], row['Max Puan']))
                    conn.commit()
                    st.success("Sınav ve ToS başarıyla buluta kaydedildi!")
                cur.close()
                conn.close()

        with tab_yonetim:
            conn = get_db_connection()
            cur = conn.cursor()
            cur.execute("SELECT sinav_kodu FROM sinavlar")
            mevcut_sinavlar = [row[0] for row in cur.fetchall()]
            if mevcut_sinavlar:
                silinecek_sinav = st.selectbox("Silinecek Sınavı Seçin", mevcut_sinavlar)
                if st.button("🗑️ Seçili Sınavı Sil", type="primary"):
                    cur.execute("DELETE FROM sinavlar WHERE sinav_kodu=%s", (silinecek_sinav,))
                    conn.commit()
                    st.success("Sınav ve bağlı tüm veriler silindi!")
                    st.rerun()
            else:
                st.info("Kayıtlı sınav bulunmuyor.")
            cur.close()
            conn.close()

        if st.session_state['rol'] == "Admin":
            with tab_kullanici:
                st.subheader("👥 Kullanıcı Yetkileri ve Şifreler")
                conn = get_db_connection()
                df_kullanicilar = pd.read_sql_query("SELECT kullanici_adi, rol, isim FROM kullanicilar", conn)
                
                with st.expander("➕ Yeni Kullanıcı Ekle veya Şifre Güncelle"):
                    c_k1, c_k2 = st.columns(2)
                    with c_k1:
                        yeni_k_adi = st.text_input("Kullanıcı Adı (ID)")
                        yeni_sifre = st.text_input("Yeni Şifre", type="password")
                    with c_k2:
                        yeni_rol = st.selectbox("Rol", ["Öğretim Elemanı", "Testing", "Admin"])
                        yeni_isim = st.text_input("Görünecek İsim (Örn: Dr. Ahmet Yılmaz)")
                    
                    if st.button("Kullanıcıyı Kaydet / Güncelle", type="primary"):
                        if yeni_k_adi and yeni_sifre and yeni_isim:
                            h_pwd = hash_sifre(yeni_sifre)
                            cur = conn.cursor()
                            cur.execute("""
                                INSERT INTO kullanicilar (kullanici_adi, sifre, rol, isim) 
                                VALUES (%s, %s, %s, %s)
                                ON CONFLICT (kullanici_adi) DO UPDATE 
                                SET sifre = EXCLUDED.sifre, rol = EXCLUDED.rol, isim = EXCLUDED.isim;
                            """, (yeni_k_adi, h_pwd, yeni_rol, yeni_isim))
                            conn.commit()
                            cur.close()
                            st.success(f"{yeni_k_adi} kullanıcısı kaydedildi!")
                            st.rerun()
                        else:
                            st.warning("Lütfen tüm alanları doldurun.")
                
                st.markdown("**Mevcut Kullanıcılar:**")
                st.dataframe(df_kullanicilar, use_container_width=True, hide_index=True)
                
                sil_k_adi = st.selectbox("Silinecek Kullanıcıyı Seçin", df_kullanicilar['kullanici_adi'].tolist() if not df_kullanicilar.empty else [])
                if st.button("Kullanıcıyı Sil") and sil_k_adi:
                    if sil_k_adi == st.session_state['kullanici_adi']:
                        st.error("Kendi hesabınızı silemezsiniz!")
                    else:
                        cur = conn.cursor()
                        cur.execute("DELETE FROM kullanicilar WHERE kullanici_adi=%s", (sil_k_adi,))
                        conn.commit()
                        cur.close()
                        st.success(f"{sil_k_adi} silindi!")
                        st.rerun()
                conn.close()
                
            with tab_atamalar:
                conn = get_db_connection()
                c_ata1, c_ata2 = st.columns(2)
                with c_ata1:
                    st.markdown("**Öğrenci Kütüğü (Master Roster)**")
                    duzenlenen_ogr = st.data_editor(pd.read_sql_query("SELECT ogrenci_no, ad_soyad, modul, sube FROM ogrenciler", conn), num_rows="dynamic", use_container_width=True)
                    if st.button("Kaydet (Öğrenciler)"):
                        cur = conn.cursor()
                        cur.execute("DELETE FROM ogrenciler")
                        for _, r in duzenlenen_ogr.iterrows():
                            cur.execute("INSERT INTO ogrenciler (ogrenci_no, ad_soyad, modul, sube) VALUES (%s, %s, %s, %s) ON CONFLICT DO NOTHING", (r['ogrenci_no'], r['ad_soyad'], r['modul'], r['sube']))
                        conn.commit()
                        cur.close()
                        st.success("Öğrenci kütüğü güncellendi!")
                with c_ata2:
                    st.markdown("**Hoca - Şube Atamaları**")
                    duzenlenen_hoca = st.data_editor(pd.read_sql_query("SELECT modul, sube, kullanici_adi FROM hoca_atamalari", conn), num_rows="dynamic", use_container_width=True)
                    if st.button("Kaydet (Öğretim Elemanları)"):
                        cur = conn.cursor()
                        cur.execute("DELETE FROM hoca_atamalari")
                        for _, r in duzenlenen_hoca.iterrows():
                            cur.execute("INSERT INTO hoca_atamalari (modul, sube, kullanici_adi) VALUES (%s, %s, %s) ON CONFLICT DO NOTHING", (r['modul'], r['sube'], r['kullanici_adi']))
                        conn.commit()
                        cur.close()
                        st.success("Atamalar güncellendi!")
                conn.close()
                
        # ----------------------------------------------------
        # KAPSAMLI ANALİTİK RAPORLAR
        # ----------------------------------------------------
        with tab_rapor:
            rapor_kapsami = st.radio("Rapor Seviyesi:", ["👤 Bireysel Öğrenci Karnesi", "🏫 Sınıf / Şube Raporu", "🌍 Modül / Kur Karşılaştırma (Benchmark)"], horizontal=True)
            st.divider()
            conn = get_db_connection()
            cur = conn.cursor()
            
            # --- 1. BİREYSEL ÖĞRENCİ KARNESİ ---
            if rapor_kapsami == "👤 Bireysel Öğrenci Karnesi":
                tum_ogrenciler = pd.read_sql_query("SELECT DISTINCT ogrenci_no, ad_soyad FROM ogrenciler ORDER BY ogrenci_no", conn)
                if not tum_ogrenciler.empty:
                    c_search, c_select = st.columns([1, 2])
                    with c_search: arama_metni = st.text_input("🔍 Öğrenci Ara (No veya İsim)").strip().lower()
                    if arama_metni:
                        filtrelenmis_ogrenciler = tum_ogrenciler[tum_ogrenciler['ogrenci_no'].str.lower().str.contains(arama_metni) | tum_ogrenciler['ad_soyad'].str.lower().str.contains(arama_metni)]
                    else: filtrelenmis_ogrenciler = tum_ogrenciler
                    
                    if not filtrelenmis_ogrenciler.empty:
                        ogrenci_etiketleri = [f"{row['ogrenci_no']} - {row['ad_soyad']}" for _, row in filtrelenmis_ogrenciler.iterrows()]
                        with c_select: secilen_etiket = st.selectbox("Öğrenci Seçimi", ogrenci_etiketleri)
                        secilen_no = secilen_etiket.split(" - ")[0]
                        
                        cur.execute("SELECT sinav_kodu, sube, veri_json FROM notlar")
                        not_kayitlari = cur.fetchall()
                        ogrenci_sinav_verisi = []
                        for skodu, s_sube, v_json in not_kayitlari:
                            df_sube = pd.read_json(io.StringIO(v_json))
                            if "Ogrenci_No" in df_sube.columns:
                                ogr_satir = df_sube[df_sube["Ogrenci_No"].astype(str) == str(secilen_no)]
                                if not ogr_satir.empty:
                                    satir_dict = ogr_satir.iloc[0].to_dict()
                                    satir_dict["Sınav Kodu"] = skodu
                                    satir_dict["Şube"] = s_sube
                                    ogrenci_sinav_verisi.append(satir_dict)
                        
                        if not ogrenci_sinav_verisi: st.warning("Bu öğrenci için girilmiş not kaydı yok.")
                        else:
                            df_gelisim_ham = pd.DataFrame(ogrenci_sinav_verisi)
                            girilen_sinav_listesi = df_gelisim_ham['Sınav Kodu'].unique().tolist()
                            rapor_turu = st.radio("Seçiniz:", ["Tek Sınavlık Karne", "Gelişim Raporu (Birden Fazla Sınav)"], horizontal=True)
                            
                            if rapor_turu == "Tek Sınavlık Karne":
                                secilen_sinav = st.selectbox("Sınavı Seçin", girilen_sinav_listesi)
                                df_rapor = df_gelisim_ham[df_gelisim_ham['Sınav Kodu'] == secilen_sinav].iloc[0]
                                df_tos = pd.read_sql_query("SELECT ana_bolum, alt_bolum, kazanimlar, max_puan FROM sinav_tos WHERE sinav_kodu=%s", conn, params=(secilen_sinav,))
                                
                                karne_satirlari = []
                                toplam_alinan, toplam_max = 0, 0
                                
                                for _, row in df_tos.iterrows():
                                    ana, alt, kazanim, mx_puan = row['ana_bolum'], row['alt_bolum'], row['kazanimlar'], row['max_puan']
                                    not_key = f"{ana} - {alt}"
                                    alinan = pd.to_numeric(df_rapor.get(not_key, 0), errors='coerce')
                                    if pd.isna(alinan) or not_key not in df_rapor:
                                        base_ana = ana.replace(' (1. Okuma)', '').replace(' (2. Okuma)', '').replace(' (3. Okuma)', '').replace(' 1st', '').replace(' 2nd', '').replace(' 3rd', '').strip()
                                        v1 = pd.to_numeric(df_rapor.get(f"{base_ana} (1. Okuma) - {alt}", pd.to_numeric(df_rapor.get(f"{base_ana} 1st - {alt}", 0), errors='coerce')), errors='coerce')
                                        v2 = pd.to_numeric(df_rapor.get(f"{base_ana} (2. Okuma) - {alt}", pd.to_numeric(df_rapor.get(f"{base_ana} 2nd - {alt}", 0), errors='coerce')), errors='coerce')
                                        v3 = pd.to_numeric(df_rapor.get(f"{base_ana} (3. Okuma) - {alt}", pd.to_numeric(df_rapor.get(f"{base_ana} 3rd - {alt}", None), errors='coerce')), errors='coerce')
                                        
                                        if pd.notnull(v3) and pd.notnull(v1) and pd.notnull(v2):
                                            d1, d2 = abs(v3 - v1), abs(v3 - v2)
                                            alinan = (v3 + v1)/2 if d1 < d2 else ((v3 + v2)/2 if d2 < d1 else (v3 + v1)/2)
                                        elif pd.notnull(v1) and pd.notnull(v2): alinan = (v1 + v2) / 2
                                        elif pd.notnull(v1): alinan = v1
                                        else: alinan = 0

                                    if pd.notnull(alinan):
                                        yuzde = (alinan / mx_puan) * 100 if mx_puan > 0 else 0
                                        clean_skill = ana.replace(' (1. Okuma)', '').replace(' (2. Okuma)', '').replace(' (3. Okuma)', '').replace(' 1st', '').replace(' 2nd', '').replace(' 3rd', '').strip()
                                        karne_satirlari.append({"Skill": clean_skill, "Section": alt, "Objective": kazanim, "Max": mx_puan, "Score": alinan, "Basari (%)": f"%{yuzde:.1f}"})
                                        toplam_alinan += alinan; toplam_max += mx_puan
                                        
                                df_detayli_karne = pd.DataFrame(karne_satirlari)
                                genel_yuzde = (toplam_alinan / toplam_max) * 100 if toplam_max > 0 else 0
                                
                                beceri_gruplari = df_detayli_karne.groupby("Skill")["Score"].sum() / df_detayli_karne.groupby("Skill")["Max"].sum() * 100
                                pdf_bar_verisi = beceri_gruplari.to_dict()
                                
                                st.divider()
                                c_yorum, c_pdf = st.columns([3, 1])
                                with c_yorum: ogretmen_yorumu = st.text_area("📝 Değerlendirme")
                                with c_pdf:
                                    st.write(""); st.write("")
                                    pdf_veri = [{'baslik': f"Academic Transcript: {secilen_etiket}", 'alt_baslik': f"Exam Code: {secilen_sinav} | Overall Score: %{genel_yuzde:.1f}", 'yorum': ogretmen_yorumu, 'beceri_grafik_verisi': pdf_bar_verisi, 'df_list': [("", df_detayli_karne)]}]
                                    st.download_button(label="📄 Akademik Karneyi İndir (PDF)", data=create_modern_pdf(pdf_veri), file_name=f"Transcript_{secilen_no}.pdf", mime="application/pdf", type="primary")
                                
                                st.markdown("### 📋 Akademik Başarı Dökümü")
                                def color_success(val):
                                    try:
                                        num = float(str(val).replace('%', ''))
                                        if num >= 75: return 'background-color: #D5F5E3; color: black;'
                                        elif num >= 50: return 'background-color: #FCF3CF; color: black;'
                                        else: return 'background-color: #FADBD8; color: black;'
                                    except: return ''
                                try: styled_df = df_detayli_karne.style.map(color_success, subset=['Basari (%)'])
                                except: styled_df = df_detayli_karne.style.applymap(color_success, subset=['Basari (%)'])
                                st.dataframe(styled_df, use_container_width=True, hide_index=True)

                            else:
                                secilen_sinavlar = st.multiselect("Dahil Edilecek Sınavlar", girilen_sinav_listesi, default=girilen_sinav_listesi)
                                df_rapor = df_gelisim_ham[df_gelisim_ham['Sınav Kodu'].isin(secilen_sinavlar)].sort_values(by="Sınav Kodu")
                                if not df_rapor.empty:
                                    st.divider()
                                    c_yorum, c_pdf = st.columns([3, 1])
                                    with c_yorum: ogretmen_yorumu = st.text_area("📝 Değerlendirme")
                                    with c_pdf:
                                        st.write(""); st.write("")
                                        pdf_veri = [{'baslik': f"Longitudinal Progress: {secilen_etiket}", 'alt_baslik': f"Multiple Exam Analysis", 'yorum': ogretmen_yorumu, 'df_list': [("Exam List", df_rapor[["Sınav Kodu", "Sınav Notu"]])], 'grafik': None}]
                                        st.download_button(label="📄 Gelişim Raporunu İndir", data=create_modern_pdf(pdf_veri), file_name=f"Progress_{secilen_no}.pdf", mime="application/pdf", type="primary")
                                    
                                    st.dataframe(df_rapor[["Sınav Kodu", "Sınav Notu", "Şube"]], use_container_width=True, hide_index=True)

            # --- 2. SINIF / ŞUBE RAPORU ---
            elif rapor_kapsami == "🏫 Sınıf / Şube Raporu":
                cur.execute("SELECT DISTINCT sube FROM notlar")
                not_kayitlari = cur.fetchall()
                if not not_kayitlari: st.info("Not kaydı bulunmuyor.")
                else:
                    tum_subeler = [r[0] for r in not_kayitlari]
                    secilen_sube = st.selectbox("Şubeyi Seçin", tum_subeler)
                    cur.execute("SELECT DISTINCT sinav_kodu FROM notlar WHERE sube=%s", (secilen_sube,))
                    sube_sinavlari = [r[0] for r in cur.fetchall()]
                    
                    st.markdown("### 🎯 Rapor Kapsamı")
                    rapor_turu_sinif = st.radio("Seçiniz:", ["Tek Sınav Sınıf Analizi", "📁 Toplu Karne Basımı (Çoklu PDF)"], horizontal=True)
                    
                    if rapor_turu_sinif == "Tek Sınav Sınıf Analizi":
                        secilen_sinav = st.selectbox("Sınav Seçin", sube_sinavlari)
                        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (secilen_sinav, secilen_sube))
                        veri_json = cur.fetchone()[0]
                        df_sinif = pd.read_json(io.StringIO(veri_json))
                        df_tos = pd.read_sql_query("SELECT ana_bolum, alt_bolum, kazanimlar, max_puan FROM sinav_tos WHERE sinav_kodu=%s", conn, params=(secilen_sinav,))
                        
                        if not df_sinif.empty and "Sınav Notu" in df_sinif.columns:
                            toplam_ogrenci = len(df_sinif)
                            sinif_ort = df_sinif["Sınav Notu"].mean()
                            medyan = df_sinif["Sınav Notu"].median()
                            max_not = df_sinif["Sınav Notu"].max()
                            min_not = df_sinif["Sınav Notu"].min()
                            
                            haric = ["Ogrenci_No", "Ad_Soyad", "Sınav Notu", "Sınav Kodu", "Şube"]
                            beceri_kolonlari = [c for c in df_sinif.columns if c not in haric and pd.to_numeric(df_sinif[c], errors='coerce').notnull().any()]
                            
                            df_liste_ui = df_sinif.copy()
                            df_liste_ui = df_liste_ui.rename(columns={"Ogrenci_No": "Student ID", "Ad_Soyad": "Name Surname", "Sınav Notu": "Total Score"})
                            secilecek_sutunlar = ["Student ID", "Name Surname"] + beceri_kolonlari + ["Total Score"]
                            df_liste_ui = df_liste_ui[secilecek_sutunlar].sort_values(by="Total Score", ascending=False)
                            
                            c_yorum, c_pdf = st.columns([3, 1])
                            with c_yorum: sinif_yorumu = st.text_area("📝 Şube Değerlendirmesi")
                            with c_pdf:
                                st.write(""); st.write("")
                                df_istatistik = pd.DataFrame({"Metric": ["Student Count", "Class Avg", "Median", "Highest", "Lowest"], "Value": [toplam_ogrenci, f"{sinif_ort:.1f}", f"{medyan:.1f}", max_not, min_not]})
                                
                                beceri_pdf_list = []
                                beceri_ui_list = []
                                testing_madde_analizi = []
                                
                                for _, r_tos in df_tos.iterrows():
                                    ana, alt, mx_puan = r_tos['ana_bolum'], r_tos['alt_bolum'], r_tos['max_puan']
                                    not_key = f"{ana} - {alt}"
                                    if not_key in df_sinif.columns:
                                        ort_puan = df_sinif[not_key].mean()
                                        yuzde = (ort_puan / mx_puan) * 100 if mx_puan > 0 else 0
                                        
                                        beceri_pdf_list.append({'isim': not_key, 'ort': ort_puan, 'max': mx_puan, 'yuzde': yuzde})
                                        beceri_ui_list.append({"Bölüm / Part": not_key, "Ortalama": f"{ort_puan:.1f} / {mx_puan}", "Başarı Oranı": yuzde})
                                        
                                        zorluk = "Kolay (P > 0.85)" if yuzde > 85 else ("Zor (P < 0.30)" if yuzde < 30 else "İdeal (0.30 - 0.85)")
                                        testing_madde_analizi.append({
                                            "Bölüm": not_key, "Ortalama": f"{ort_puan:.1f} / {mx_puan}", 
                                            "Zorluk İndeksi (P)": f"{yuzde/100:.2f}", "Pedagojik Durum": zorluk
                                        })
                                
                                pdf_veri_yatay = {'baslik': f"Class Analysis Report - {secilen_sube}", 'alt_baslik': f"Exam: {secilen_sinav}", 'yorum': sinif_yorumu, 'df_kpi': df_istatistik, 'beceri_listesi': beceri_pdf_list, 'df_liste': df_liste_ui}
                                st.download_button(label="📄 Sınıf Raporunu İndir (Landscape PDF)", data=create_landscape_class_pdf(pdf_veri_yatay), file_name=f"{secilen_sube}_Class_Report.pdf", mime="application/pdf", type="primary")

                            st.divider()
                            st.markdown("**📊 Sınıf Özet Panosu (Class Overview)**")
                            m1, m2, m3, m4, m5 = st.columns(5)
                            m1.metric("Student Count", toplam_ogrenci)
                            m2.metric("Class Avg", f"{sinif_ort:.1f}")
                            m3.metric("Median", f"{medyan:.1f}")
                            m4.metric("Highest Score", max_not)
                            m5.metric("Lowest Score", min_not)
                            
                            st.markdown("**📈 Beceri / Bölüm Başarı Oranları**")
                            if beceri_ui_list:
                                df_beceri_ui = pd.DataFrame(beceri_ui_list)
                                st.dataframe(df_beceri_ui, use_container_width=True, hide_index=True, column_config={"Başarı Oranı": st.column_config.ProgressColumn("Başarı Oranı (%)", min_value=0, max_value=100, format="%f%%")})

                            st.markdown("**📝 Detaylı Öğrenci Listesi**")
                            st.dataframe(df_liste_ui, use_container_width=True, hide_index=True)
                            
                            st.divider()
                            with st.expander("🔬 Testing Ofisi & Admin Özel: Madde & Kazanım Zorluk Analitiği (Item Facility Index)", expanded=False):
                                st.caption("Bu bölüm sınavın ölçme-değerlendirme kalitesini izlemek için Testing Ofisi'ne ayrılmıştır; öğretmen raporlarında veya PDF çıktısında yer almaz.")
                                if testing_madde_analizi:
                                    st.dataframe(pd.DataFrame(testing_madde_analizi), use_container_width=True, hide_index=True)

                    elif rapor_turu_sinif == "📁 Toplu Karne Basımı (Çoklu PDF)":
                        secilen_sinav = st.selectbox("Karnesi Basılacak Sınavı Seçin", sube_sinavlari)
                        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (secilen_sinav, secilen_sube))
                        veri_json = cur.fetchone()[0]
                        df_sinif = pd.read_json(io.StringIO(veri_json))
                        df_tos = pd.read_sql_query("SELECT ana_bolum, alt_bolum, kazanimlar, max_puan FROM sinav_tos WHERE sinav_kodu=%s", conn, params=(secilen_sinav,))
                        
                        ogrenci_listesi = [f"{row['Ogrenci_No']} - {row['Ad_Soyad']}" for _, row in df_sinif.iterrows()]
                        secilen_ogrenciler = st.multiselect("Öğrencileri Seçin", ogrenci_listesi, default=ogrenci_listesi)
                        
                        if st.button("🖨️ Toplu Oluştur", type="primary") and secilen_ogrenciler:
                            rapor_verileri = []
                            for ogr_etiket in secilen_ogrenciler:
                                ogr_no = ogr_etiket.split(" - ")[0]
                                df_ogr_tek = df_sinif[df_sinif["Ogrenci_No"].astype(str) == str(ogr_no)].iloc[0]
                                
                                karne_satirlari = []
                                t_alinan, t_max = 0, 0
                                for _, r_tos in df_tos.iterrows():
                                    ana, alt, kazanim, mx_puan = r_tos['ana_bolum'], r_tos['alt_bolum'], r_tos['max_puan']
                                    not_key = f"{ana} - {alt}"
                                    alinan = pd.to_numeric(df_ogr_tek.get(not_key, 0), errors='coerce')
                                    if pd.isna(alinan) or not_key not in df_ogr_tek:
                                        base_ana = ana.replace(' (1. Okuma)', '').replace(' (2. Okuma)', '').replace(' (3. Okuma)', '').replace(' 1st', '').replace(' 2nd', '').replace(' 3rd', '').strip()
                                        v1 = pd.to_numeric(df_ogr_tek.get(f"{base_ana} (1. Okuma) - {alt}", pd.to_numeric(df_ogr_tek.get(f"{base_ana} 1st - {alt}", 0), errors='coerce')), errors='coerce')
                                        v2 = pd.to_numeric(df_ogr_tek.get(f"{base_ana} (2. Okuma) - {alt}", pd.to_numeric(df_ogr_tek.get(f"{base_ana} 2nd - {alt}", 0), errors='coerce')), errors='coerce')
                                        v3 = pd.to_numeric(df_ogr_tek.get(f"{base_ana} (3. Okuma) - {alt}", pd.to_numeric(df_ogr_tek.get(f"{base_ana} 3rd - {alt}", None), errors='coerce')), errors='coerce')
                                        if pd.notnull(v3) and pd.notnull(v1) and pd.notnull(v2):
                                            d1, d2 = abs(v3 - v1), abs(v3 - v2)
                                            alinan = (v3 + v1)/2 if d1 < d2 else ((v3 + v2)/2 if d2 < d1 else (v3 + v1)/2)
                                        elif pd.notnull(v1) and pd.notnull(v2): alinan = (v1 + v2) / 2
                                        elif pd.notnull(v1): alinan = v1
                                        else: alinan = 0
                                    
                                    if pd.notnull(alinan):
                                        yuzde = (alinan / mx_puan) * 100 if mx_puan > 0 else 0
                                        clean_skill = ana.replace(' (1. Okuma)', '').replace(' (2. Okuma)', '').replace(' (3. Okuma)', '').replace(' 1st', '').replace(' 2nd', '').replace(' 3rd', '').strip()
                                        karne_satirlari.append({"Skill": clean_skill, "Section": alt, "Objective": kazanim, "Max": mx_puan, "Score": alinan, "Basari (%)": f"%{yuzde:.1f}"})
                                        t_alinan += alinan; t_max += mx_puan
                                
                                genel_yuzde = (t_alinan / t_max) * 100 if t_max > 0 else 0
                                df_kisi = pd.DataFrame(karne_satirlari)
                                beceri_gruplari = df_kisi.groupby("Skill")["Score"].sum() / df_kisi.groupby("Skill")["Max"].sum() * 100
                                
                                rapor_verileri.append({
                                    'baslik': f"Academic Transcript: {ogr_etiket}",
                                    'alt_baslik': f"Exam: {secilen_sinav} | Class: {secilen_sube} | Overall: %{genel_yuzde:.1f}",
                                    'yorum': "", 'beceri_grafik_verisi': beceri_gruplari.to_dict(), 'df_list': [("Score Details", df_kisi)]
                                })
                                
                            st.download_button(label="📥 Toplu Karneleri İndir (Tek PDF)", data=create_modern_pdf(rapor_verileri), file_name=f"{secilen_sube}_Toplu_Karneler.pdf", mime="application/pdf", type="primary")

            # --- 3. MODÜL / KUR KARŞILAŞTIRMA (BENCHMARK) ---
            elif rapor_kapsami == "🌍 Modül / Kur Karşılaştırma (Benchmark)":
                st.markdown("### 🌍 Sınav Bazlı Şubeler Arası Karşılaştırma")
                cur.execute("SELECT sinav_kodu FROM sinavlar")
                acik_sinavlar = [row[0] for row in cur.fetchall()]
                if not acik_sinavlar: st.info("Aktif sınav bulunmuyor.")
                else:
                    benchmark_sinav = st.selectbox("Analiz Edilecek Sınavı Seçin", acik_sinavlar)
                    cur.execute("SELECT sube, veri_json FROM notlar WHERE sinav_kodu=%s", (benchmark_sinav,))
                    ilgili_notlar = cur.fetchall()
                    if not ilgili_notlar: st.warning("Bu sınav için henüz şube notu girilmemiş.")
                    else:
                        tum_subeler_df = pd.DataFrame()
                        for s_sube, v_json in ilgili_notlar:
                            df_temp = pd.read_json(io.StringIO(v_json))
                            df_temp["Şube"] = s_sube
                            tum_subeler_df = pd.concat([tum_subeler_df, df_temp], ignore_index=True)
                        
                        if "Sınav Notu" in tum_subeler_df.columns:
                            sube_ortalamalari = tum_subeler_df.groupby("Şube")["Sınav Notu"].mean().reset_index().sort_values(by="Sınav Notu", ascending=False)
                            
                            c1, c2 = st.columns(2)
                            with c1:
                                st.markdown("**📊 Şubeler Arası Ortalama Kıyaslaması**")
                                fig_bar = px.bar(sube_ortalamalari, x="Şube", y="Sınav Notu", text_auto='.1f', color="Sınav Notu", color_continuous_scale="Viridis")
                                st.plotly_chart(fig_bar, use_container_width=True)
                            with c2:
                                st.markdown("**📦 Şubelerin Not Dağılımı (Box Plot)**")
                                fig_box = px.box(tum_subeler_df, x="Şube", y="Sınav Notu", color="Şube")
                                st.plotly_chart(fig_box, use_container_width=True)
                                
            cur.close()
            conn.close()

    # ------------------------------------------
    # ÖĞRETİM ELEMANI PANELİ (VERİ GİRİŞİ VE OBS AKTARIM)
    # ------------------------------------------
    elif st.session_state['rol'] == "Öğretim Elemanı":
        st.subheader("📝 Hızlı Veri Giriş Terminali")
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("SELECT sinav_kodu FROM sinavlar")
        acik_sinavlar = [row[0] for row in cur.fetchall()]
        if not acik_sinavlar: st.warning("Aktif sınav bulunmuyor."); st.stop()
        secilen_islem_sinavi = st.selectbox("Aktif Sınavlar", acik_sinavlar)
        ilgili_modul = secilen_islem_sinavi.split("_")[0]
        cur.execute("SELECT sube FROM hoca_atamalari WHERE modul=%s AND kullanici_adi=%s", (ilgili_modul, st.session_state['kullanici_adi']))
        hoca_subeleri = [row[0] for row in cur.fetchall()]
        if not hoca_subeleri: st.error("Bu modülde adınıza atanmış şube bulunmuyor."); st.stop()
        secilen_sube = st.selectbox("Atandığınız Şubeler", hoca_subeleri)
        rubrik_df = pd.read_sql_query("SELECT * FROM sinav_tos WHERE sinav_kodu=%s", conn, params=(secilen_islem_sinavi,))
        cur.execute("SELECT tolerans FROM sinav_ayarlari WHERE sinav_kodu=%s", (secilen_islem_sinavi,))
        tolerans_val = cur.fetchone()
        sinav_toleransi = tolerans_val[0] if tolerans_val else 15.0
        
        st.markdown(f"### 📊 {secilen_sube} - Not Tablosu")
        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (secilen_islem_sinavi, secilen_sube))
        kaydedilmis_veri = cur.fetchone()
        
        if kaydedilmis_veri:
            df_giris = pd.read_json(io.StringIO(kaydedilmis_veri[0]))
            st.success("Daha önce buluta kaydettiğiniz veriler yüklendi.")
        else:
            gercek_ogrenciler = pd.read_sql_query("SELECT ogrenci_no AS \"Ogrenci_No\", ad_soyad AS \"Ad_Soyad\" FROM ogrenciler WHERE modul=%s AND sube=%s", conn, params=(ilgili_modul, secilen_sube))
            df_giris = gercek_ogrenciler.copy() if not gercek_ogrenciler.empty else pd.DataFrame(columns=["Ogrenci_No", "Ad_Soyad"])
            for _, row in rubrik_df.iterrows(): df_giris[f"{row['ana_bolum']} - {row['alt_bolum']}"] = 0.0 
            df_giris["Sınav Notu"] = 0.0 
            
        girilen_notlar_df = st.data_editor(df_giris, num_rows="dynamic", use_container_width=True, column_config={"Sınav Notu": st.column_config.NumberColumn("Sınav Notu", disabled=True)})
        
        if st.button("💾 Notları Buluta Kaydet", type="primary"):
            hatalar = []
            max_puan_sozlugu = {f"{r['ana_bolum']} - {r['alt_bolum']}": float(r['max_puan']) for _, r in rubrik_df.iterrows()}
            for idx, row in girilen_notlar_df.iterrows():
                for col, m_puan in max_puan_sozlugu.items():
                    val = pd.to_numeric(row.get(col), errors='coerce')
                    if pd.notnull(val) and val > m_puan: hatalar.append(f"**{row.get('Ogrenci_No', '')}**: '{col}' notu ({val}), max ({m_puan}) aşıyor!")
            if hatalar: st.error("🛑 HATA!"); [st.warning(h) for h in hatalar]; st.stop()
            toplam_notlar = []
            
            uyarilar = []
            for col in girilen_notlar_df.columns:
                karsilik_col = None
                if "(1. Okuma)" in col: karsilik_col = col.replace("(1. Okuma)", "(2. Okuma)")
                elif " 1st" in col: karsilik_col = col.replace(" 1st", " 2nd")
                
                if karsilik_col and karsilik_col in girilen_notlar_df.columns:
                    for idx, row in girilen_notlar_df.iterrows():
                        not1, not2 = pd.to_numeric(row.get(col), errors='coerce'), pd.to_numeric(row.get(karsilik_col), errors='coerce')
                        if pd.notnull(not1) and pd.notnull(not2):
                            fark = abs(not1 - not2)
                            if fark >= sinav_toleransi:
                                uyarilar.append(f"**{row['Ogrenci_No']} - {row['Ad_Soyad']}** | Fark: **{fark}** (Sınır: {sinav_toleransi}). 3. Okuma Gerekli!")
                                
            for idx, row in girilen_notlar_df.iterrows():
                ogrenci_toplam = 0.0
                islenen_temeller = set()
                for col in max_puan_sozlugu.keys():
                    base_col = col.replace("(1. Okuma)", "").replace("(2. Okuma)", "").replace("(3. Okuma)", "").replace(" 1st", "").replace(" 2nd", "").replace(" 3rd", "").strip()
                    if base_col in islenen_temeller: continue
                    
                    if col == base_col and col in girilen_notlar_df.columns:
                        val = pd.to_numeric(row.get(col), errors='coerce')
                        if pd.notnull(val): ogrenci_toplam += val
                        islenen_temeller.add(base_col)
                    else:
                        ilgili_kolonlar = [k for k in max_puan_sozlugu.keys() if k.replace("(1. Okuma)", "").replace("(2. Okuma)", "").replace("(3. Okuma)", "").replace(" 1st", "").replace(" 2nd", "").replace(" 3rd", "").strip() == base_col]
                        v1, v2, v3 = None, None, None
                        for k in ilgili_kolonlar:
                            v = pd.to_numeric(row.get(k), errors='coerce')
                            if pd.notnull(v):
                                if "1. Okuma" in k or "1st" in k: v1 = v
                                elif "2. Okuma" in k or "2nd" in k: v2 = v
                                elif "3. Okuma" in k or "3rd" in k: v3 = v
                                else: v1 = v
                        
                        final_part_score = 0
                        if v1 is not None and v2 is not None and v3 is not None:
                            d1, d2 = abs(v3 - v1), abs(v3 - v2)
                            final_part_score = (v3 + v1)/2 if d1 < d2 else ((v3 + v2)/2 if d2 < d1 else (v3 + v1)/2)
                        elif v1 is not None and v2 is not None: final_part_score = (v1 + v2) / 2
                        elif v1 is not None: final_part_score = v1
                        elif v2 is not None: final_part_score = v2
                        elif v3 is not None: final_part_score = v3

                        ogrenci_toplam += final_part_score
                        islenen_temeller.add(base_col)
                toplam_notlar.append(round(ogrenci_toplam, 2))
            girilen_notlar_df["Sınav Notu"] = toplam_notlar
            
            cur.execute("""
                INSERT INTO notlar (sinav_kodu, sube, veri_json) 
                VALUES (%s, %s, %s)
                ON CONFLICT (sinav_kodu, sube) DO UPDATE 
                SET veri_json = EXCLUDED.veri_json;
            """, (secilen_islem_sinavi, secilen_sube, girilen_notlar_df.to_json(orient="records")))
            conn.commit()
            cur.close()
            conn.close()
            st.success("Notlar bulut veritabanına başarıyla kaydedildi!")
            if uyarilar:
                st.error("🚨 ÇİFT OKUMA ÇAKIŞMA UYARISI:")
                for uyari in uyarilar: st.warning(uyari)
                
            st.rerun()
            
        st.divider()
        with st.expander("📤 OBS Not Aktarım Ekranı (Kopyala-Yapıştır)"):
            st.info("Aşağıdaki tablodan 'Öğrenci No' ve 'Sınav Notu' sütunlarını seçerek OBS'ye doğrudan yapıştırabilirsiniz.")
            if not girilen_notlar_df.empty:
                df_obs = girilen_notlar_df[["Ogrenci_No", "Ad_Soyad", "Sınav Notu"]].copy()
                st.dataframe(df_obs, use_container_width=True, hide_index=True)
                
                obs_text = ""
                for _, r in df_obs.iterrows(): obs_text += f"{r['Ogrenci_No']}\t{r['Sınav Notu']}\n"
                st.text_area("Hızlı Kopyalama Alanı (Öğr No ve Not)", obs_text, height=150)
            else:
                st.warning("Gösterilecek veri yok.")